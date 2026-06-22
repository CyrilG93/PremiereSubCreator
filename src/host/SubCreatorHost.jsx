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

function subcreator_decode_base64_to_binary_string(input) {
  // // Decode base64 text into a binary-safe ExtendScript string so template Source Text payloads can be restored in host.
  var sanitized = String(input || "").replace(/\s+/g, "");
  if (!sanitized) {
    return "";
  }

  var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  var buffer = 0;
  var bits = 0;
  var bytes = [];

  for (var index = 0; index < sanitized.length; index += 1) {
    var character = sanitized.charAt(index);
    if (character === "=") {
      break;
    }

    var alphabetIndex = alphabet.indexOf(character);
    if (alphabetIndex < 0) {
      continue;
    }

    buffer = (buffer << 6) | alphabetIndex;
    bits += 6;

    while (bits >= 8) {
      bits -= 8;
      bytes.push((buffer >> bits) & 255);
    }
  }

  return subcreator_byte_array_to_binary_string(bytes);
}

function subcreator_ok(data) {
  // // Normalize successful host responses for panel-side parsing.
  return JSON.stringify({ ok: true, data: data });
}

function subcreator_error(message, debug) {
  // // Normalize failure responses for panel-side parsing.
  var payload = { ok: false, error: String(message) };
  if (debug !== undefined) {
    payload.debug = debug;
  }
  return JSON.stringify(payload);
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

function subcreator_get_sequence_identity(sequence) {
  // // Capture stable project and sequence identifiers so panel-side timing metadata can be keyed safely.
  var project = app && app.project ? app.project : null;
  var projectDocumentId = "";
  var projectPath = "";
  var sequenceID = "";
  var sequenceName = "";

  try {
    projectDocumentId = subcreator_trim_string(project && project.documentID ? String(project.documentID) : "");
  } catch (error) {
    projectDocumentId = "";
  }

  try {
    projectPath = subcreator_trim_string(project && project.path ? String(project.path) : "");
  } catch (error) {
    projectPath = "";
  }

  try {
    sequenceID = subcreator_trim_string(sequence && sequence.sequenceID ? String(sequence.sequenceID) : "");
  } catch (error) {
    sequenceID = "";
  }

  try {
    sequenceName = subcreator_trim_string(sequence && sequence.name ? String(sequence.name) : "");
  } catch (error) {
    sequenceName = "";
  }

  return {
    projectDocumentId: projectDocumentId,
    projectPath: projectPath,
    sequenceID: sequenceID,
    sequenceName: sequenceName
  };
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
    var selected = null;
    if (/windows/i.test(String($.os || ""))) {
      // // Windows CEP treats function filters as literal strings, so use a native filter string there.
      selected = File.openDialog("Select SRT subtitle file", "SRT subtitle files:*.srt");
    } else {
      // // macOS supports callback filters, which lets folders stay navigable while restricting files to `.srt`.
      selected = File.openDialog("Select SRT subtitle file", function (candidate) {
        if (candidate instanceof Folder) {
          return true;
        }
        return /\.srt$/i.test(String(candidate.name || ""));
      });
    }
    if (!selected) {
      return subcreator_ok({ path: "" });
    }

    return subcreator_ok({ path: selected.fsName });
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_pick_corrected_transcript_file() {
  // // Open native picker for corrected transcript sources used by WhisperX alignment.
  try {
    var selected = null;
    if (/windows/i.test(String($.os || ""))) {
      // // Windows CEP needs a filter string so Explorer shows only supported corrected-transcript files.
      selected = File.openDialog("Select corrected transcript file", "Corrected transcript files:*.srt;*.txt");
    } else {
      // // macOS can keep folder navigation while filtering to `.srt` and `.txt` transcript files.
      selected = File.openDialog("Select corrected transcript file", function (candidate) {
        if (candidate instanceof Folder) {
          return true;
        }
        return /\.(srt|txt)$/i.test(String(candidate.name || ""));
      });
    }
    if (!selected) {
      return subcreator_ok({ path: "" });
    }

    return subcreator_ok({ path: selected.fsName });
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_normalize_system_path(value) {
  // // Normalize file-system paths for the current host platform before passing them into export APIs.
  var normalized = String(value || "");
  return subcreator_is_windows() ? normalized.replace(/\//g, "\\") : normalized.replace(/\\/g, "/");
}

function subcreator_read_sequence_in_out_range(sequence) {
  // // Centralize active-sequence In/Out extraction so SRT, Whisper, and corrected-align use the same Premiere range lookup.
  var rangeStartSeconds = NaN;
  var rangeEndSeconds = NaN;

  try {
    rangeStartSeconds = subcreator_to_seconds(
      typeof sequence.getInPointAsTime === "function"
        ? sequence.getInPointAsTime()
        : typeof sequence.getInPoint === "function"
          ? sequence.getInPoint()
          : null
    );
  } catch (inPointError) {
    rangeStartSeconds = NaN;
  }

  try {
    rangeEndSeconds = subcreator_to_seconds(
      typeof sequence.getOutPointAsTime === "function"
        ? sequence.getOutPointAsTime()
        : typeof sequence.getOutPoint === "function"
          ? sequence.getOutPoint()
          : null
    );
  } catch (outPointError) {
    rangeEndSeconds = NaN;
  }

  if ((!isFinite(rangeStartSeconds) || rangeStartSeconds < 0) && isFinite(rangeEndSeconds) && rangeEndSeconds > 0) {
    rangeStartSeconds = 0;
  }

  if (!isFinite(rangeStartSeconds) || !isFinite(rangeEndSeconds) || rangeStartSeconds < 0 || rangeEndSeconds <= rangeStartSeconds) {
    rangeStartSeconds = NaN;
    rangeEndSeconds = NaN;
  }

  return {
    rangeStartSeconds: isFinite(rangeStartSeconds) ? Number(rangeStartSeconds) : null,
    rangeEndSeconds: isFinite(rangeEndSeconds) ? Number(rangeEndSeconds) : null
  };
}

function subcreator_get_active_sequence_range() {
  // // Expose the current sequence In/Out range to CEP so non-audio sources can respect the same user-selected range.
  try {
    if (!app || !app.project || !app.project.activeSequence) {
      return subcreator_error("No active sequence in Premiere.");
    }

    var sequence = app.project.activeSequence;
    var range = subcreator_read_sequence_in_out_range(sequence);
    return subcreator_ok({
      rangeStartSeconds: range.rangeStartSeconds,
      rangeEndSeconds: range.rangeEndSeconds,
      sequenceName: String(sequence.name || "")
    });
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_build_audio_preset_candidates() {
  // // Build likely Adobe system-preset roots across installed major versions.
  var candidates = [];
  var currentYear = new Date().getFullYear() + 1;

  if (subcreator_is_windows()) {
    var roots = [];
    try {
      subcreator_runtime_push_unique(roots, $.getenv("ProgramFiles"));
    } catch (error) {}
    try {
      subcreator_runtime_push_unique(roots, $.getenv("ProgramFiles(x86)"));
    } catch (error2) {}

    for (var rootIndex = 0; rootIndex < roots.length; rootIndex += 1) {
      var root = String(roots[rootIndex] || "");
      for (var year = currentYear; year >= 2023; year -= 1) {
        subcreator_runtime_push_unique(
          candidates,
          root + "/Adobe/Adobe Media Encoder " + year + "/MediaIO/systempresets"
        );
        subcreator_runtime_push_unique(
          candidates,
          root + "/Adobe/Adobe Premiere Pro " + year + "/MediaIO/systempresets"
        );
      }
      subcreator_runtime_push_unique(candidates, root + "/Adobe/Adobe Media Encoder Beta/MediaIO/systempresets");
      subcreator_runtime_push_unique(candidates, root + "/Adobe/Adobe Media Encoder (Beta)/MediaIO/systempresets");
    }

    return candidates;
  }

  for (var macYear = currentYear; macYear >= 2023; macYear -= 1) {
    var mediaEncoderAppName = "Adobe Media Encoder " + macYear;
    var premiereAppName = "Adobe Premiere Pro " + macYear;
    subcreator_runtime_push_unique(
      candidates,
      "/Applications/" + mediaEncoderAppName + "/" + mediaEncoderAppName + ".app/Contents/MediaIO/systempresets"
    );
    subcreator_runtime_push_unique(candidates, "/Applications/" + mediaEncoderAppName + ".app/Contents/MediaIO/systempresets");
    subcreator_runtime_push_unique(
      candidates,
      "/Applications/" + premiereAppName + "/" + premiereAppName + ".app/Contents/MediaIO/systempresets"
    );
    subcreator_runtime_push_unique(candidates, "/Applications/" + premiereAppName + ".app/Contents/MediaIO/systempresets");
  }

  subcreator_runtime_push_unique(
    candidates,
    "/Applications/Adobe Media Encoder (Beta)/Adobe Media Encoder (Beta).app/Contents/MediaIO/systempresets"
  );
  subcreator_runtime_push_unique(candidates, "/Applications/Adobe Media Encoder (Beta).app/Contents/MediaIO/systempresets");
  subcreator_runtime_push_unique(candidates, "/Applications/Adobe Media Encoder Beta/Adobe Media Encoder Beta.app/Contents/MediaIO/systempresets");
  subcreator_runtime_push_unique(candidates, "/Applications/Adobe Media Encoder Beta.app/Contents/MediaIO/systempresets");

  return candidates;
}

function subcreator_find_audio_preset_in_systempresets(rootPath) {
  // // Scan Adobe Media Encoder system presets for a stable WAV audio-only preset across versions.
  var rootFolder = new Folder(subcreator_normalize_system_path(rootPath || ""));
  if (!rootFolder.exists) {
    return "";
  }

  var preferredPatterns = [/^Waveform Audio 48kHz 16-bit\.epr$/i, /^WAV 48kHz 16 bit\.epr$/i];

  function scanFolder(folderRef) {
    if (!folderRef || !folderRef.exists) {
      return "";
    }

    var entries = folderRef.getFiles();
    for (var entryIndex = 0; entryIndex < entries.length; entryIndex += 1) {
      var entry = entries[entryIndex];
      if (entry instanceof File) {
        var entryName = String(entry.name || "");
        for (var patternIndex = 0; patternIndex < preferredPatterns.length; patternIndex += 1) {
          if (preferredPatterns[patternIndex].test(entryName)) {
            return entry.fsName;
          }
        }
      }
    }

    for (var folderIndex = 0; folderIndex < entries.length; folderIndex += 1) {
      var child = entries[folderIndex];
      if (child instanceof Folder) {
        var match = scanFolder(child);
        if (match) {
          return match;
        }
      }
    }

    return "";
  }

  return scanFolder(rootFolder);
}

function subcreator_find_audio_export_preset(preferredPath) {
  // // Resolve a usable audio-only export preset so active-sequence Whisper export stays automatic.
  var normalizedPreferred = subcreator_trim_string(preferredPath || "");
  if (normalizedPreferred) {
    var preferredFile = new File(subcreator_normalize_system_path(normalizedPreferred));
    if (preferredFile.exists) {
      return preferredFile.fsName;
    }
  }

  var candidates = subcreator_build_audio_preset_candidates();
  for (var i = 0; i < candidates.length; i += 1) {
    var candidatePath = subcreator_find_audio_preset_in_systempresets(candidates[i]);
    if (candidatePath) {
      return candidatePath;
    }
  }

  return "";
}

function subcreator_wait_for_file_stable(filePath, timeoutMs, stablePasses) {
  // // Poll an exported file until it exists with a non-zero size and stops growing for a few checks.
  var normalizedPath = subcreator_normalize_system_path(filePath);
  var deadline = new Date().getTime() + Math.max(Number(timeoutMs) || 0, 1000);
  var lastSize = -1;
  var stableCount = 0;

  while (new Date().getTime() < deadline) {
    var fileRef = new File(normalizedPath);
    if (fileRef.exists && Number(fileRef.length) > 0) {
      var currentSize = Number(fileRef.length);
      if (currentSize === lastSize) {
        stableCount += 1;
        if (stableCount >= Math.max(Number(stablePasses) || 0, 2)) {
          return true;
        }
      } else {
        lastSize = currentSize;
        stableCount = 0;
      }
    }

    $.sleep(500);
  }

  var finalFile = new File(normalizedPath);
  return finalFile.exists && Number(finalFile.length) > 0;
}

function subcreator_file_debug_snapshot(filePath) {
  // // Capture file state from ExtendScript so failed exports show whether Premiere touched the WAV.
  try {
    var normalizedPath = subcreator_normalize_system_path(filePath || "");
    if (!normalizedPath) {
      return { path: "", exists: false };
    }

    var fileRef = new File(normalizedPath);
    if (!fileRef.exists) {
      return { path: normalizedPath, exists: false };
    }

    return {
      path: fileRef.fsName,
      exists: true,
      length: Number(fileRef.length || 0),
      modified: String(fileRef.modified || "")
    };
  } catch (error) {
    return { path: String(filePath || ""), exists: "unknown", error: String(error) };
  }
}

function subcreator_sequence_debug_snapshot(sequence) {
  // // Capture stable sequence facts without depending on optional Premiere APIs.
  var debug = {
    name: "",
    sequenceID: "",
    videoTracks: "",
    audioTracks: "",
    frameSize: "",
    timebase: ""
  };

  try {
    debug.name = String(sequence && sequence.name ? sequence.name : "");
  } catch (nameError) {}
  try {
    debug.sequenceID = String(sequence && sequence.sequenceID ? sequence.sequenceID : "");
  } catch (idError) {}
  try {
    debug.videoTracks = String(sequence && sequence.videoTracks ? sequence.videoTracks.numTracks : "");
  } catch (videoTrackError) {}
  try {
    debug.audioTracks = String(sequence && sequence.audioTracks ? sequence.audioTracks.numTracks : "");
  } catch (audioTrackError) {}
  try {
    debug.frameSize = String(sequence.frameSizeHorizontal || "") + "x" + String(sequence.frameSizeVertical || "");
  } catch (frameSizeError) {}
  try {
    debug.timebase = String(sequence.timebase || "");
  } catch (timebaseError) {}

  return debug;
}

function subcreator_try_encoder_sequence_export(sequence, outputPath, presetPath, workAreaType, exportErrors) {
  // // Queue the sequence through Adobe Media Encoder only when the isolated Premiere export has failed.
  if (!app.encoder || typeof app.encoder.encodeSequence !== "function") {
    exportErrors.push("encodeSequence unavailable");
    return false;
  }

  try {
    if (typeof app.encoder.launchEncoder === "function") {
      app.encoder.launchEncoder();
    }
  } catch (launchError) {
    exportErrors.push("launchEncoder: " + launchError);
  }

  try {
    var jobId = app.encoder.encodeSequence(sequence, outputPath, presetPath, workAreaType, 0);
    if (jobId && String(jobId) !== "0") {
      if (typeof app.encoder.startBatch === "function") {
        app.encoder.startBatch();
      }
      return true;
    }

    exportErrors.push("encodeSequence returned no job id");
  } catch (encoderError) {
    exportErrors.push("encodeSequence: " + encoderError);
  }

  return false;
}

function subcreator_try_direct_sequence_export(sequence, outputPath, presetPath, workAreaType, exportErrors) {
  // // Render inside Premiere; the CEP panel isolates this call because Premiere 26.x can return an opaque evalScript error.
  if (typeof sequence.exportAsMediaDirect !== "function") {
    exportErrors.push("exportAsMediaDirect unavailable");
    return false;
  }

  try {
    var directResult = sequence.exportAsMediaDirect(outputPath, presetPath, workAreaType);
    if (directResult) {
      return true;
    }

    exportErrors.push("exportAsMediaDirect returned false");
  } catch (directExportError) {
    exportErrors.push("exportAsMediaDirect: " + directExportError);
  }

  return false;
}

function subcreator_get_sequence_export_capabilities() {
  // // Report export API availability without launching an export, keeping failure diagnostics low-risk.
  try {
    var sequence = app && app.project ? app.project.activeSequence : null;
    return subcreator_ok({
      hostOs: String($.os || ""),
      hasActiveSequence: Boolean(sequence),
      sequence: sequence ? subcreator_sequence_debug_snapshot(sequence) : {},
      hasEncoder: Boolean(app && app.encoder),
      hasLaunchEncoder: Boolean(app && app.encoder && typeof app.encoder.launchEncoder === "function"),
      hasEncodeSequence: Boolean(app && app.encoder && typeof app.encoder.encodeSequence === "function"),
      hasExportAsMediaDirect: Boolean(sequence && typeof sequence.exportAsMediaDirect === "function")
    });
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_export_active_sequence_audio(payloadEncoded) {
  // // Render the active-sequence audible mix to a temporary WAV file for Whisper transcription.
  var debug = {
    hostOs: String($.os || ""),
    payloadDecoded: false,
    exportMode: "",
    rangeMode: "",
    workAreaType: "",
    requestedOutputPath: "",
    normalizedOutputPath: "",
    requestedPresetPath: "",
    resolvedPresetPath: "",
    sequence: {},
    activeRange: {},
    exportErrors: [],
    fileBefore: {},
    fileAfter: {}
  };

  try {
    var payloadText = subcreator_decode_payload(payloadEncoded || "");
    var payload = payloadText ? JSON.parse(payloadText) : {};
    debug.payloadDecoded = true;
    debug.rangeMode = String(payload.rangeMode || "");
    debug.exportMode = String(payload.exportMode || "premiere_direct");
    debug.requestedOutputPath = String(payload.outputPath || "");
    debug.requestedPresetPath = String(payload.presetPath || "");

    if (!app || !app.project || !app.project.activeSequence) {
      return subcreator_error("No active sequence available for Whisper export.", debug);
    }

    var sequence = app.project.activeSequence;
    debug.sequence = subcreator_sequence_debug_snapshot(sequence);
    var outputPath = subcreator_normalize_system_path(payload.outputPath || "");
    debug.normalizedOutputPath = outputPath;
    if (!outputPath) {
      return subcreator_error("Missing output path for Whisper sequence export.", debug);
    }

    var presetPath = subcreator_find_audio_export_preset(payload.presetPath || "");
    debug.resolvedPresetPath = presetPath;
    if (!presetPath) {
      return subcreator_error("Unable to locate the WAV preset for Whisper sequence export.", debug);
    }

    var requestedInOut = String(payload.rangeMode || "") === "in_out";
    var activeRange = requestedInOut ? subcreator_read_sequence_in_out_range(sequence) : { rangeStartSeconds: null, rangeEndSeconds: null };
    debug.activeRange = activeRange;
    var hasValidRange =
      requestedInOut &&
      activeRange &&
      activeRange.rangeStartSeconds !== null &&
      activeRange.rangeEndSeconds !== null &&
      isFinite(activeRange.rangeStartSeconds) &&
      isFinite(activeRange.rangeEndSeconds) &&
      activeRange.rangeEndSeconds > activeRange.rangeStartSeconds;
    // // Fall back to the full sequence when the panel requested In/Out but the active sequence has no valid range.
    var workAreaType = hasValidRange ? 1 : 0;
    debug.workAreaType = String(workAreaType);
    if (!hasValidRange) {
      activeRange = { rangeStartSeconds: null, rangeEndSeconds: null };
    }

    var outputFile = new File(outputPath);
    var outputFolder = outputFile.parent;
    if (outputFolder && !outputFolder.exists) {
      outputFolder.create();
    }

    var exportMode = String(payload.exportMode || "premiere_direct");
    var exportErrors = [];
    debug.fileBefore = subcreator_file_debug_snapshot(outputPath);
    // // Keep each exporter in a separate evalScript call so a Premiere direct-export failure cannot prevent the AME fallback.
    var exportTriggered =
      exportMode === "media_encoder"
        ? subcreator_try_encoder_sequence_export(sequence, outputPath, presetPath, workAreaType, exportErrors)
        : subcreator_try_direct_sequence_export(sequence, outputPath, presetPath, workAreaType, exportErrors);
    debug.exportErrors = exportErrors;
    debug.fileAfter = subcreator_file_debug_snapshot(outputPath);

    if (!exportTriggered) {
      return subcreator_error("Unable to start active-sequence audio export. " + exportErrors.join(" | "), debug);
    }

    if (!subcreator_wait_for_file_stable(outputPath, 600000, 3)) {
      debug.fileAfter = subcreator_file_debug_snapshot(outputPath);
      return subcreator_error("Timed out waiting for exported Whisper audio file: " + outputPath, debug);
    }
    debug.fileAfter = subcreator_file_debug_snapshot(outputPath);

    return subcreator_ok({
      audioPath: outputFile.fsName,
      presetPath: presetPath,
      exportMethod: exportMode,
      sequenceName: String(sequence.name || ""),
      rangeStartSeconds: activeRange.rangeStartSeconds,
      rangeEndSeconds: activeRange.rangeEndSeconds,
      debug: debug
    });
  } catch (error) {
    debug.exception = String(error);
    return subcreator_error(error, debug);
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
      // // Strip the UTF-8 BOM emitted by Windows PowerShell 5.1 before parsing older runtime configs.
      payload = payload.replace(/^\uFEFF/, "");
      var parsed = JSON.parse(payload);
      if (parsed && typeof parsed === "object") {
        parsed.__sourcePath = candidatePath;
        return parsed;
      }
    } catch (error) {}
  }

  if (subcreator_is_windows()) {
    // // Recover the standard private runtime even when the installer config is absent or unreadable.
    var localAppData = "";
    try {
      localAppData = subcreator_trim_string($.getenv("LOCALAPPDATA"));
    } catch (error) {}

    if (localAppData) {
      var runtimeRoot = localAppData + "/SubCreator/runtime";
      var pythonPath = runtimeRoot + "/python/python.exe";
      var whisperPath = runtimeRoot + "/python/Scripts/whisper.exe";
      var ffmpegPath = runtimeRoot + "/ffmpeg/bin/ffmpeg.exe";
      if (new File(pythonPath).exists) {
        return {
          pythonCommand: pythonPath,
          pythonPath: pythonPath,
          whisperPath: new File(whisperPath).exists ? whisperPath : "",
          ffmpegPath: new File(ffmpegPath).exists ? ffmpegPath : "",
          pathHints: [
            runtimeRoot + "/python",
            runtimeRoot + "/python/Scripts",
            runtimeRoot + "/ffmpeg/bin",
            "C:/Windows/System32"
          ],
          __sourcePath: runtimeRoot + " (automatic recovery)"
        };
      }
    }
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
      " --output_format all --output_dir " +
      subcreator_quote_cmd(outputDir) +
      " --fp16 False --word_timestamps True";

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
    " --output_format all --output_dir " +
    subcreator_quote_posix(outputDir) +
    " --fp16 False --word_timestamps True";

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

function subcreator_find_whisper_json_file(tempFolder, baseName) {
  // // Resolve the JSON file created by Whisper so panel can reuse exact word timestamps.
  var directPath = new File(tempFolder.fsName + "/" + baseName + ".json");
  if (directPath.exists) {
    return directPath;
  }

  var files = tempFolder.getFiles("*.json");
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
      return subcreator_error("Host system.callSystem indisponible. Active le mode Node CEP pour Whisper.");
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

    var jsonText = "";
    var jsonFile = subcreator_find_whisper_json_file(tempFolder, baseName);
    if (jsonFile && jsonFile.exists) {
      jsonText = subcreator_read_file_text(jsonFile);
    }

    return subcreator_ok({
      srtText: srtText,
      jsonText: jsonText,
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
  var components = subcreator_get_mogrt_components_from_track_item(trackItem);
  return components.length > 0 ? components[0] : null;
}

function subcreator_get_mogrt_components_from_track_item(trackItem) {
  // // Collect every usable Essential Graphics component because Premiere-authored MOGRTs do not always expose text on the first component.
  if (!trackItem) {
    return [];
  }

  var components = [];

  function rememberComponent(candidate) {
    if (!candidate || !candidate.properties || typeof candidate.properties.numItems !== "number") {
      return;
    }

    for (var index = 0; index < components.length; index += 1) {
      if (components[index] === candidate) {
        return;
      }
    }

    components.push(candidate);
  }

  try {
    if (typeof trackItem.getMGTComponent === "function") {
      rememberComponent(trackItem.getMGTComponent());
    }
  } catch (mgtError) {}

  if (trackItem.components && typeof trackItem.components.numItems === "number" && trackItem.components.numItems > 0) {
    for (var componentIndex = 0; componentIndex < trackItem.components.numItems; componentIndex += 1) {
      rememberComponent(trackItem.components[componentIndex]);
    }
  }

  return components;
}

function subcreator_collect_selected_mogrt_items(sequence) {
  // // Collect selected timeline clips that expose a valid MOGRT component.
  var selectedItems = subcreator_get_selected_track_items(sequence);
  var mogrtItems = [];

  for (var index = 0; index < selectedItems.length; index += 1) {
    var trackItem = selectedItems[index];
    if (subcreator_get_mogrt_component_from_track_item(trackItem)) {
      mogrtItems.push(trackItem);
    }
  }

  return mogrtItems;
}

function subcreator_sort_track_items_by_time(items) {
  // // Keep subtitle text operations deterministic by sorting selected MOGRT clips in timeline order.
  var sortedItems = (items || []).slice(0);
  sortedItems.sort(function (left, right) {
    var leftStart = subcreator_to_seconds(left && (left.start || left.inPoint || left.startTime));
    var rightStart = subcreator_to_seconds(right && (right.start || right.inPoint || right.startTime));
    if (leftStart < rightStart) {
      return -1;
    }
    if (leftStart > rightStart) {
      return 1;
    }

    var leftEnd = subcreator_to_seconds(left && (left.end || left.outPoint || left.endTime));
    var rightEnd = subcreator_to_seconds(right && (right.end || right.outPoint || right.endTime));
    if (leftEnd < rightEnd) {
      return -1;
    }
    if (leftEnd > rightEnd) {
      return 1;
    }
    return 0;
  });
  return sortedItems;
}

function subcreator_find_track_item_video_track_index(sequence, trackItem) {
  // // Resolve the owning video track index so rebuilt text clips can stay on the same track.
  if (!sequence || !sequence.videoTracks || typeof sequence.videoTracks.numTracks !== "number" || !trackItem) {
    return -1;
  }

  try {
    if (trackItem.parentTrack) {
      for (var parentTrackIndex = 0; parentTrackIndex < sequence.videoTracks.numTracks; parentTrackIndex += 1) {
        if (sequence.videoTracks[parentTrackIndex] === trackItem.parentTrack) {
          return parentTrackIndex;
        }
      }
    }
  } catch (parentTrackError) {}

  function safeLowerText(value) {
    return subcreator_trim_string(String(value || "")).toLowerCase();
  }

  function sameRoundedTime(leftValue, rightValue) {
    var leftSeconds = subcreator_to_seconds(leftValue);
    var rightSeconds = subcreator_to_seconds(rightValue);
    if (isNaN(leftSeconds) || isNaN(rightSeconds)) {
      return false;
    }
    return Math.abs(leftSeconds - rightSeconds) <= 0.002;
  }

  function itemsLookEquivalent(leftItem, rightItem) {
    if (!leftItem || !rightItem) {
      return false;
    }
    if (leftItem === rightItem) {
      return true;
    }

    var sameStart = sameRoundedTime(leftItem.start || leftItem.inPoint || leftItem.startTime, rightItem.start || rightItem.inPoint || rightItem.startTime);
    var sameEnd = sameRoundedTime(leftItem.end || leftItem.outPoint || leftItem.endTime, rightItem.end || rightItem.outPoint || rightItem.endTime);
    if (!sameStart || !sameEnd) {
      return false;
    }

    var leftProjectItemName = safeLowerText(leftItem.projectItem && leftItem.projectItem.name);
    var rightProjectItemName = safeLowerText(rightItem.projectItem && rightItem.projectItem.name);
    if (leftProjectItemName && rightProjectItemName && leftProjectItemName === rightProjectItemName) {
      return true;
    }

    var leftName = safeLowerText(leftItem.name);
    var rightName = safeLowerText(rightItem.name);
    if (leftName && rightName && leftName === rightName) {
      return true;
    }

    var leftText = safeLowerText(subcreator_extract_text_from_mogrt_item(leftItem));
    var rightText = safeLowerText(subcreator_extract_text_from_mogrt_item(rightItem));
    if (leftText && rightText && leftText === rightText) {
      return true;
    }

    if (leftProjectItemName && rightName && leftProjectItemName === rightName) {
      return true;
    }

    if (leftName && rightProjectItemName && leftName === rightProjectItemName) {
      return true;
    }

    return false;
  }

  for (var trackIndex = 0; trackIndex < sequence.videoTracks.numTracks; trackIndex += 1) {
    var track = sequence.videoTracks[trackIndex];
    var clips = subcreator_collection_to_array(track ? track.clips : null);
    for (var clipIndex = 0; clipIndex < clips.length; clipIndex += 1) {
      if (itemsLookEquivalent(clips[clipIndex], trackItem)) {
        return trackIndex;
      }
    }
  }

  return -1;
}

function subcreator_collect_resolved_selected_mogrt_items(sequence) {
  // // Drop selected MOGRT references that Premiere can no longer map to one real track item after rebuild operations.
  var rawSelection = subcreator_sort_track_items_by_time(subcreator_collect_selected_mogrt_items(sequence));
  var resolvedSelection = [];

  for (var index = 0; index < rawSelection.length; index += 1) {
    if (subcreator_find_track_item_video_track_index(sequence, rawSelection[index]) >= 0) {
      resolvedSelection.push(rawSelection[index]);
    }
  }

  return resolvedSelection;
}

function subcreator_build_selected_mogrt_text_signature(sequence, trackItems) {
  // // Encode enough selection details so text-apply can reject stale selections after the user changes clips.
  var sortedItems = subcreator_sort_track_items_by_time(trackItems || []);
  var parts = [];

  for (var index = 0; index < sortedItems.length; index += 1) {
    var trackItem = sortedItems[index];
    var videoTrackIndex = subcreator_find_track_item_video_track_index(sequence, trackItem);
    var startSeconds = subcreator_to_seconds(trackItem && (trackItem.start || trackItem.inPoint || trackItem.startTime));
    var endSeconds = subcreator_to_seconds(trackItem && (trackItem.end || trackItem.outPoint || trackItem.endTime));
    var clipText = subcreator_trim_string(String(subcreator_extract_text_from_mogrt_item(trackItem) || "").replace(/\s+/g, " "));
    parts.push(
      [
        String(videoTrackIndex),
        String(Math.round(Number(startSeconds || 0) * 1000)),
        String(Math.round(Number(endSeconds || 0) * 1000)),
        clipText
      ].join("|")
    );
  }

  return parts.join("||");
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
    key.indexOf("colorize") !== -1 ||
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

function subcreator_visual_is_crop_label(displayName) {
  // // Premiere crop controls are percentage sliders and should stay in a 0..100 range.
  var key = String(displayName || "").toLowerCase();
  var normalizedKey = subcreator_visual_normalize_label_key(displayName);
  return (
    key.indexOf("crop") !== -1 ||
    normalizedKey.indexOf("crop") !== -1 ||
    normalizedKey.indexOf("recadr") !== -1
  );
}

function subcreator_visual_normalize_label_key(label) {
  // // Normalize labels for robust matching across accents/typos/localized variants.
  return String(label || "")
    .toLowerCase()
    .replace(/[àáâãäå]/g, "a")
    .replace(/[èéêë]/g, "e")
    .replace(/[ìíîï]/g, "i")
    .replace(/[òóôõö]/g, "o")
    .replace(/[ùúûü]/g, "u")
    .replace(/ç/g, "c")
    .replace(/[^a-z0-9]/g, "");
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
  var normalizedKey = subcreator_visual_normalize_label_key(displayName);
  var numericValue = subcreator_visual_to_number(rawValue);

  if (isNaN(numericValue)) {
    numericValue = 0;
  }

  if (
    key.indexOf("opacity") !== -1 ||
    key.indexOf("opacite") !== -1 ||
    normalizedKey.indexOf("opacity") !== -1 ||
    normalizedKey.indexOf("opacite") !== -1 ||
    normalizedKey.indexOf("ocapite") !== -1
  ) {
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

  if (subcreator_visual_is_crop_label(displayName)) {
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

  var normalizedKey = subcreator_visual_normalize_label_key(displayName);
  var isOpacityLike =
    String(displayName || "").toLowerCase().indexOf("opacity") !== -1 ||
    String(displayName || "").toLowerCase().indexOf("opacite") !== -1 ||
    normalizedKey.indexOf("opacity") !== -1 ||
    normalizedKey.indexOf("opacite") !== -1 ||
    normalizedKey.indexOf("ocapite") !== -1;
  if (isOpacityLike) {
    // // Some MOGRT metadata reports oversized max values for opacity while UI is effectively clamped to 100.
    minValue = 0;
    maxValue = 100;
    if (isNaN(stepValue) || stepValue <= 0) {
      stepValue = 1;
    }
  }

  if (subcreator_visual_is_crop_label(displayName)) {
    // // Premiere can expose misleading crop metadata like `-50..50`, but the UI is really `0..100`.
    minValue = 0;
    maxValue = 100;
    if (isNaN(stepValue) || stepValue <= 0) {
      stepValue = 1;
    }
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

function subcreator_visual_build_select_options(displayName, rawValue, groupPath) {
  // // Build known dropdown option sets for menu-like numeric controls.
  var key = String(displayName || "").toLowerCase();
  var groupKey = String(groupPath || "").toLowerCase();
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

  function buildEnumeratedOptions(values, labels) {
    var options = [];
    for (var optionIndex = 0; optionIndex < values.length && optionIndex < labels.length; optionIndex += 1) {
      options.push({
        value: values[optionIndex],
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

  if (key.indexOf("highlight mode") !== -1) {
    // // Keep AE dropdowns readable in the Visual Editor even when CEP exposes them as plain numeric controls.
    if (numericValue >= 0 && numericValue <= 1) {
      return buildLabeledRange(0, ["Progressive", "Cumulative"]);
    }
    if (numericValue >= 1 && numericValue <= 2) {
      return buildLabeledRange(1, ["Progressive", "Cumulative"]);
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

  if (key.indexOf("blend mode") !== -1 || key === "blendmode") {
    if (groupKey.indexOf("opacity") !== -1) {
      // // Premiere's main clip blend mode behaves like Adobe's classic internal blend-mode enum values rather than a 0-based list.
      return buildEnumeratedOptions(
        [
          12, 13, 15, 16, 17, 18, 19,
          20, 22, 23, 24, 25,
          26, 27, 28, 29, 30, 31, 32,
          33, 34, 35, 36,
          37, 38, 39, 40
        ],
        [
          "Normal", "Dissolve", "Darken", "Multiply", "Color Burn", "Linear Burn", "Darker Color",
          "Lighten", "Screen", "Color Dodge", "Linear Dodge (Add)", "Lighter Color",
          "Overlay", "Soft Light", "Hard Light", "Vivid Light", "Linear Light", "Pin Light", "Hard Mix",
          "Difference", "Exclusion", "Subtract", "Divide",
          "Hue", "Saturation", "Color", "Luminosity"
        ]
      );
    }

    if (groupKey.indexOf("wonder glow") !== -1 || groupKey.indexOf("glow") !== -1) {
      if (numericValue >= 0 && numericValue <= 1) {
        return buildLabeledRange(0, ["Screen", "Additive"]);
      }
      if (numericValue >= 1 && numericValue <= 2) {
        return buildLabeledRange(1, ["Screen", "Additive"]);
      }
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
var subcreator_visual_text_style_option_cache = {
  families: {},
  styles: {}
};

function subcreator_visual_reset_group_sequence_axis_preferences() {
  // // Reset per-group vector scaling hints before reading a new selection.
  subcreator_visual_group_sequence_axis_preferences = {};
}

function subcreator_visual_reset_text_style_option_cache() {
  // // Reset discovered font/style options before reading a new selection.
  subcreator_visual_text_style_option_cache = {
    families: {},
    styles: {}
  };
}

function subcreator_visual_register_text_style_option(optionType, value) {
  // // Keep a global deduplicated cache of font families/styles found in current selection.
  var normalizedType = subcreator_trim_string(String(optionType || "")).toLowerCase();
  var text = subcreator_trim_string(String(value || ""));
  if (!text) {
    return;
  }

  if (normalizedType === "family") {
    subcreator_visual_text_style_option_cache.families[text.toLowerCase()] = text;
  } else if (normalizedType === "style") {
    subcreator_visual_text_style_option_cache.styles[text.toLowerCase()] = text;
  }
}

function subcreator_visual_read_text_style_option_cache(optionType) {
  // // Read sorted options from the current read-session cache.
  var normalizedType = subcreator_trim_string(String(optionType || "")).toLowerCase();
  var source = normalizedType === "style" ? subcreator_visual_text_style_option_cache.styles : subcreator_visual_text_style_option_cache.families;
  var result = [];

  for (var key in source) {
    if (source.hasOwnProperty(key)) {
      result.push(source[key]);
    }
  }

  result.sort(function (left, right) {
    var a = String(left || "").toLowerCase();
    var b = String(right || "").toLowerCase();
    if (a < b) {
      return -1;
    }
    if (a > b) {
      return 1;
    }
    return 0;
  });

  return result;
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

function subcreator_visual_should_force_graphic_parameters_axis_scale(displayName, groupPath) {
  // // Some AE Essential Graphics vectors stay normalized even when Premiere displays them as pixel X/Y values in Properties.
  var displayKey = String(displayName || "").toLowerCase();
  var groupKey = String(groupPath || "").toLowerCase();
  if (groupKey.indexOf("graphic parameters") === -1) {
    return false;
  }

  if (groupKey.indexOf("/ box controls") !== -1 && (displayKey.indexOf("padding") !== -1 || displayKey.indexOf("offset") !== -1)) {
    return true;
  }

  if (groupKey.indexOf("/ offset") !== -1 && displayKey.indexOf("offset") !== -1) {
    return true;
  }

  return false;
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
  var looksLikeAnchor = displayKey.indexOf("anchor") !== -1;
  var shouldForceGraphicParametersAxisScale = subcreator_visual_should_force_graphic_parameters_axis_scale(displayName, groupPath);

  if (shouldForceGraphicParametersAxisScale) {
    // // Keep AE Graphic Parameters vectors like Box Padding/Offset aligned with the pixel units shown in Premiere Properties.
    return {
      mode: vectorMode,
      scale: [width, height, 1, 1].slice(0, vectorValues.length),
      candidateId: "graphic_parameters_axis",
      score: 0
    };
  }

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
    if (
      looksLikeAnchor &&
      subcreator_visual_group_prefers_sequence_axis(groupPath) &&
      subcreator_visual_vector_looks_normalized_position(vectorValues)
    ) {
      // // When Position already proved this group uses normalized sequence coordinates, Anchor Point should follow the same axis conversion.
      return {
        mode: vectorMode,
        scale: [width, height, 1, 1].slice(0, vectorValues.length),
        candidateId: "anchor_group_sequence_axis",
        score: 0
      };
    }

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

function subcreator_visual_round_number_for_display(value, decimals) {
  // // Round panel-facing numeric values to keep visual editor compact and readable.
  var numericValue = Number(value);
  if (isNaN(numericValue)) {
    return 0;
  }
  var digits = Math.max(Number(decimals || 1), 0);
  var factor = Math.pow(10, digits);
  var rounded = Math.round(numericValue * factor) / factor;
  if (Math.abs(rounded) < 0.0000001) {
    return 0;
  }
  return rounded;
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

function subcreator_visual_is_generic_font_family_key(normalizedKey) {
  // // Catch additional template-specific family keys (for example `fontEditValue`).
  return (
    normalizedKey === "fonteditvalue" ||
    normalizedKey === "mfonteditvalue" ||
    normalizedKey === "fontfamilyvalue" ||
    normalizedKey === "fontfamilyeditvalue" ||
    normalizedKey === "mfontfamilyvalue" ||
    normalizedKey === "mfontfamilyeditvalue"
  );
}

function subcreator_visual_is_generic_font_style_key(normalizedKey) {
  // // Catch additional template-specific style keys (for example `fontStyleEditValue`).
  return (
    normalizedKey === "fontstylevalue" ||
    normalizedKey === "fontstyleeditvalue" ||
    normalizedKey === "mfontstylevalue" ||
    normalizedKey === "mfontstyleeditvalue" ||
    normalizedKey === "fontfauxstylevalue" ||
    normalizedKey === "fontfauxstyleeditvalue"
  );
}

function subcreator_visual_is_generic_font_size_key(normalizedKey) {
  // // Catch additional template-specific size keys (for example `fontSizeEditValue`).
  return (
    normalizedKey === "fontsizevalue" ||
    normalizedKey === "fontsizeeditvalue" ||
    normalizedKey === "mfontsizevalue" ||
    normalizedKey === "mfontsizeeditvalue"
  );
}

function subcreator_visual_extract_first_string(value) {
  // // Extract first textual item from scalar or array payloads.
  if (typeof value === "string") {
    var direct = subcreator_trim_string(value);
    return direct ? direct : "";
  }

  if (value && typeof value.length === "number" && value.length > 0) {
    for (var index = 0; index < value.length; index += 1) {
      if (typeof value[index] === "string") {
        var item = subcreator_trim_string(String(value[index] || ""));
        if (item) {
          return item;
        }
      }
    }
  }

  return "";
}

function subcreator_visual_extract_first_number(value) {
  // // Extract first numeric item from scalar or array payloads (ignoring booleans).
  if (typeof value === "number") {
    return Number(value);
  }

  if (typeof value === "string") {
    var fromString = Number(value);
    return isNaN(fromString) ? NaN : fromString;
  }

  if (typeof value === "boolean") {
    return NaN;
  }

  if (value && typeof value.length === "number" && value.length > 0) {
    for (var index = 0; index < value.length; index += 1) {
      if (typeof value[index] === "boolean") {
        continue;
      }
      var itemNumber = Number(value[index]);
      if (!isNaN(itemNumber)) {
        return itemNumber;
      }
    }
  }

  return NaN;
}

function subcreator_visual_split_font_token(token) {
  // // Split font token strings like `Montserrat-Bold` or `Avenir Regular` into family/style parts.
  var text = subcreator_trim_string(String(token || "")).replace(/_/g, " ").replace(/\s+/g, " ");
  if (!text) {
    return { family: "", style: "" };
  }

  var hyphenIndex = text.indexOf("-");
  if (hyphenIndex >= 1) {
    var family = subcreator_trim_string(text.substring(0, hyphenIndex));
    var style = subcreator_trim_string(text.substring(hyphenIndex + 1));
    return {
      family: family || text,
      style: style
    };
  }

  var styleKeywords = {
    thin: true,
    hairline: true,
    extralight: true,
    ultralight: true,
    light: true,
    book: true,
    regular: true,
    roman: true,
    plain: true,
    medium: true,
    semibold: true,
    demibold: true,
    bold: true,
    extrabold: true,
    ultrabold: true,
    black: true,
    heavy: true,
    italic: true,
    oblique: true,
    condensed: true,
    narrow: true,
    expanded: true,
    extended: true,
    display: true,
    caps: true,
    smallcaps: true
  };
  var words = text.split(/\s+/);
  if (words.length > 1) {
    var styleWords = [];
    var cursor = words.length - 1;
    while (cursor >= 0) {
      var probe = String(words[cursor] || "").toLowerCase();
      if (!styleKeywords[probe]) {
        break;
      }
      styleWords.unshift(words[cursor]);
      cursor -= 1;
    }
    if (styleWords.length > 0 && cursor >= 0) {
      return {
        family: subcreator_trim_string(words.slice(0, cursor + 1).join(" ")),
        style: subcreator_trim_string(styleWords.join(" "))
      };
    }
  }

  return { family: text, style: "" };
}

function subcreator_visual_join_font_token(family, style, fallbackToken) {
  // // Rebuild font token while preserving a sane fallback when style is empty.
  var normalizedFamily = subcreator_trim_string(String(family || ""));
  var normalizedStyle = subcreator_trim_string(String(style || ""));
  if (!normalizedFamily) {
    var fallback = subcreator_trim_string(String(fallbackToken || ""));
    return fallback;
  }
  if (!normalizedStyle) {
    return normalizedFamily;
  }
  return normalizedFamily + "-" + normalizedStyle;
}

function subcreator_visual_compact_font_token_part(value) {
  // // Build compact token fragments used by many internal font ids (`AvenirNext-DemiBold`).
  return subcreator_trim_string(String(value || "")).replace(/[^A-Za-z0-9]+/g, "");
}

function subcreator_visual_push_unique_string(target, value) {
  // // Append one unique string while preserving insertion order.
  if (!target || typeof target.length !== "number") {
    return;
  }
  var text = subcreator_trim_string(String(value || ""));
  if (!text) {
    return;
  }
  for (var index = 0; index < target.length; index += 1) {
    if (String(target[index] || "").toLowerCase() === text.toLowerCase()) {
      return;
    }
  }
  target.push(text);
}

function subcreator_visual_has_own_entries(target) {
  // // ExtendScript-safe replacement for `Object.keys(target).length > 0`.
  if (!target || typeof target !== "object") {
    return false;
  }
  for (var key in target) {
    if (target.hasOwnProperty(key)) {
      return true;
    }
  }
  return false;
}

function subcreator_visual_list_font_style_aliases(style) {
  // // Expand display style aliases so `Regular`, `Roman` and `Plain` all resolve.
  var normalizedStyle = subcreator_trim_string(String(style || ""));
  var aliases = [];
  if (!normalizedStyle) {
    aliases.push("");
    return aliases;
  }
  subcreator_visual_push_unique_string(aliases, normalizedStyle);
  var styleKey = normalizedStyle.toLowerCase();
  if (styleKey === "regular" || styleKey === "roman" || styleKey === "plain") {
    subcreator_visual_push_unique_string(aliases, "Regular");
    subcreator_visual_push_unique_string(aliases, "Roman");
    subcreator_visual_push_unique_string(aliases, "Plain");
  }
  return aliases;
}

function subcreator_visual_build_font_token_candidates(family, style, providedToken, fallbackToken) {
  // // Generate token variants because Premiere accepts different internal separators across fonts.
  var candidates = [];
  var requestedFamily = subcreator_trim_string(String(family || ""));
  var requestedStyle = subcreator_trim_string(String(style || ""));
  var exactToken = subcreator_trim_string(String(providedToken || ""));
  var fallbackParts = subcreator_visual_split_font_token(fallbackToken);
  var exactParts = subcreator_visual_split_font_token(exactToken);
  var familyVariants = [];
  var styleVariants = [];
  var exactTokenLooksCanonical = !!(exactToken && (exactToken.indexOf(" ") < 0 || !exactParts.style));

  if (exactTokenLooksCanonical) {
    // // Prefer trusted exact tokens (`Amarillo`, `Avenir-Regular`) before rebuilding variants.
    subcreator_visual_push_unique_string(candidates, exactToken);
  }

  subcreator_visual_push_unique_string(familyVariants, requestedFamily || exactParts.family || fallbackParts.family);
  subcreator_visual_push_unique_string(familyVariants, exactParts.family);
  subcreator_visual_push_unique_string(familyVariants, fallbackParts.family);

  var primaryFamily = familyVariants.length > 0 ? familyVariants[0] : "";
  var compactPrimaryFamily = subcreator_visual_compact_font_token_part(primaryFamily);
  if (compactPrimaryFamily && compactPrimaryFamily.toLowerCase() !== primaryFamily.toLowerCase()) {
    subcreator_visual_push_unique_string(familyVariants, compactPrimaryFamily);
  }

  var resolvedStyle = requestedStyle || exactParts.style || fallbackParts.style || "Regular";
  var resolvedStyleAliases = subcreator_visual_list_font_style_aliases(resolvedStyle);
  for (var styleAliasIndex = 0; styleAliasIndex < resolvedStyleAliases.length; styleAliasIndex += 1) {
    var aliasValue = resolvedStyleAliases[styleAliasIndex];
    subcreator_visual_push_unique_string(styleVariants, aliasValue);
    var compactAlias = subcreator_visual_compact_font_token_part(aliasValue);
    if (compactAlias && compactAlias.toLowerCase() !== aliasValue.toLowerCase()) {
      subcreator_visual_push_unique_string(styleVariants, compactAlias);
    }
  }

  for (var familyIndex = 0; familyIndex < familyVariants.length; familyIndex += 1) {
    var familyVariant = familyVariants[familyIndex];
    if (!familyVariant) {
      continue;
    }
    for (var styleIndex = 0; styleIndex < styleVariants.length; styleIndex += 1) {
      var styleVariant = styleVariants[styleIndex];
      if (!styleVariant) {
        continue;
      }
      subcreator_visual_push_unique_string(candidates, subcreator_visual_join_font_token(familyVariant, styleVariant, fallbackToken));
      subcreator_visual_push_unique_string(candidates, familyVariant + " " + styleVariant);
    }
    subcreator_visual_push_unique_string(candidates, familyVariant);
  }

  if (exactToken && !exactTokenLooksCanonical) {
    // // Filename-like aliases with spaces (`Avenir Regular`) are tried after canonical variants.
    subcreator_visual_push_unique_string(candidates, exactToken);
  }

  return candidates;
}

function subcreator_visual_normalize_font_compare_key(value) {
  // // Build comparison key for font family matching across display/internal naming variants.
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function subcreator_visual_font_family_matches_value(rawValue, expectedFamily) {
  // // Check whether readback payload resolves to requested font family.
  var expectedKey = subcreator_visual_normalize_font_compare_key(expectedFamily);
  if (!expectedKey) {
    return null;
  }

  var extracted = subcreator_visual_extract_text_style_from_value(rawValue);
  if (extracted && extracted.fontFamily) {
    var extractedKey = subcreator_visual_normalize_font_compare_key(extracted.fontFamily);
    if (!extractedKey) {
      return null;
    }
    return extractedKey === expectedKey;
  }

  if (typeof rawValue === "string") {
    var rawKey = subcreator_visual_normalize_font_compare_key(rawValue);
    if (rawKey && rawKey.indexOf(expectedKey) !== -1) {
      return true;
    }
  }

  return null;
}

function subcreator_visual_push_unique_option(list, value) {
  // // Push one option text value without duplicates.
  if (!list) {
    return;
  }
  var text = subcreator_trim_string(String(value || ""));
  if (!text) {
    return;
  }
  var key = text.toLowerCase();
  for (var index = 0; index < list.length; index += 1) {
    if (String(list[index] || "").toLowerCase() === key) {
      return;
    }
  }
  list.push(text);
}

function subcreator_visual_extract_string_options(value) {
  // // Extract candidate string options from scalar/array payload values.
  var result = [];
  if (typeof value === "string") {
    subcreator_visual_push_unique_option(result, value);
    return result;
  }

  if (value && typeof value.length === "number" && value.length > 0) {
    for (var index = 0; index < value.length; index += 1) {
      if (typeof value[index] === "string") {
        subcreator_visual_push_unique_option(result, value[index]);
      }
    }
  }

  return result;
}

function subcreator_visual_extract_first_boolean(value) {
  // // Extract boolean-like values from scalar/array payloads.
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "number") {
    if (value === 0 || value === 1) {
      return value === 1;
    }
    return null;
  }

  if (typeof value === "string") {
    var normalized = subcreator_trim_string(value).toLowerCase();
    if (normalized === "true" || normalized === "1") {
      return true;
    }
    if (normalized === "false" || normalized === "0") {
      return false;
    }
    return null;
  }

  if (value && typeof value.length === "number" && value.length > 0) {
    for (var index = 0; index < value.length; index += 1) {
      var fromItem = subcreator_visual_extract_first_boolean(value[index]);
      if (typeof fromItem === "boolean") {
        return fromItem;
      }
    }
  }

  return null;
}

function subcreator_visual_register_style_for_family(styleMap, family, style) {
  // // Register one style entry under a font family without duplicates.
  if (!styleMap || typeof styleMap !== "object") {
    return;
  }

  var normalizedFamily = subcreator_trim_string(String(family || ""));
  var normalizedStyle = subcreator_trim_string(String(style || ""));
  if (!normalizedFamily || !normalizedStyle) {
    return;
  }

  var targetFamilyKey = "";
  var normalizedFamilyKey = normalizedFamily.toLowerCase();
  for (var existingFamilyKey in styleMap) {
    if (!styleMap.hasOwnProperty(existingFamilyKey)) {
      continue;
    }
    if (String(existingFamilyKey || "").toLowerCase() === normalizedFamilyKey) {
      targetFamilyKey = existingFamilyKey;
      break;
    }
  }
  if (!targetFamilyKey) {
    targetFamilyKey = normalizedFamily;
  }

  if (!styleMap[targetFamilyKey] || typeof styleMap[targetFamilyKey].length !== "number") {
    styleMap[targetFamilyKey] = [];
  }

  var bucket = styleMap[targetFamilyKey];
  subcreator_visual_push_unique_option(bucket, normalizedStyle);
}

function subcreator_visual_build_style_map_output(styleMap) {
  // // Export style-map cache using display family names with sorted style lists.
  var output = {};
  if (!styleMap || typeof styleMap !== "object") {
    return output;
  }

  for (var familyKey in styleMap) {
    if (!styleMap.hasOwnProperty(familyKey)) {
      continue;
    }
    var familyBucket = styleMap[familyKey];
    if (!familyBucket || typeof familyBucket.length !== "number" || familyBucket.length === 0) {
      continue;
    }
    var styles = [];
    for (var styleIndex = 0; styleIndex < familyBucket.length; styleIndex += 1) {
      subcreator_visual_push_unique_option(styles, familyBucket[styleIndex]);
    }
    if (!styles.length) {
      continue;
    }
    styles.sort(function (left, right) {
      var a = String(left || "").toLowerCase();
      var b = String(right || "").toLowerCase();
      if (a < b) {
        return -1;
      }
      if (a > b) {
        return 1;
      }
      return 0;
    });
    output[String(familyKey)] = styles;
  }

  return output;
}

function subcreator_visual_is_font_flag_key(normalizedKey, flagKey) {
  // // Detect boolean font-style toggle keys in text payloads.
  var key = String(normalizedKey || "");
  var flag = String(flagKey || "");
  if (!key || !flag) {
    return false;
  }

  if (key === flag + "value" || key === flag + "editvalue" || key === "m" + flag + "value" || key === "m" + flag + "editvalue") {
    return true;
  }

  return key.indexOf(flag) !== -1;
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
    fontToken: "",
    fontSize: NaN,
    fontFamilyOptions: [],
    fontStyleOptions: [],
    fontStylesByFamily: {},
    fontFauxStyleEditable: null,
    fontFsBold: null,
    fontFsItalic: null,
    fontFsAllCaps: null,
    fontFsSmallCaps: null
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

      if (!result.fontFamily && subcreator_visual_is_generic_font_family_key(normalizedKey)) {
        var fontToken = subcreator_visual_extract_first_string(value);
        if (fontToken) {
          if (!result.fontToken) {
            result.fontToken = fontToken;
          }
          var tokenParts = subcreator_visual_split_font_token(fontToken);
          if (tokenParts.family) {
            result.fontFamily = tokenParts.family;
          }
          if (!result.fontStyle && tokenParts.style) {
            result.fontStyle = tokenParts.style;
          }
        }
      }
      if (subcreator_visual_is_generic_font_family_key(normalizedKey)) {
        var tokenOptions = subcreator_visual_extract_string_options(value);
        for (var tokenIndex = 0; tokenIndex < tokenOptions.length; tokenIndex += 1) {
          var tokenPartsOption = subcreator_visual_split_font_token(tokenOptions[tokenIndex]);
          subcreator_visual_push_unique_option(result.fontFamilyOptions, tokenPartsOption.family);
          if (tokenPartsOption.style) {
            subcreator_visual_push_unique_option(result.fontStyleOptions, tokenPartsOption.style);
            subcreator_visual_register_style_for_family(result.fontStylesByFamily, tokenPartsOption.family, tokenPartsOption.style);
          }
        }
      }

      if (!result.fontFamily && subcreator_visual_is_font_family_key(normalizedKey)) {
        var familyValue = subcreator_visual_extract_first_string(value);
        if (familyValue) {
          result.fontFamily = familyValue;
        }
      }
      if (subcreator_visual_is_font_family_key(normalizedKey)) {
        var familyOptions = subcreator_visual_extract_string_options(value);
        for (var familyIndex = 0; familyIndex < familyOptions.length; familyIndex += 1) {
          subcreator_visual_push_unique_option(result.fontFamilyOptions, familyOptions[familyIndex]);
        }
      }

      if (!result.fontStyle && subcreator_visual_is_font_style_key(normalizedKey)) {
        var styleValue = subcreator_visual_extract_first_string(value);
        if (styleValue) {
          result.fontStyle = styleValue;
        }
      }
      if (subcreator_visual_is_font_style_key(normalizedKey)) {
        var styleOptions = subcreator_visual_extract_string_options(value);
        for (var styleIndex = 0; styleIndex < styleOptions.length; styleIndex += 1) {
          subcreator_visual_push_unique_option(result.fontStyleOptions, styleOptions[styleIndex]);
          if (result.fontFamily) {
            subcreator_visual_register_style_for_family(result.fontStylesByFamily, result.fontFamily, styleOptions[styleIndex]);
          }
        }
      }
      if (!result.fontStyle && subcreator_visual_is_generic_font_style_key(normalizedKey)) {
        var genericStyleValue = subcreator_visual_extract_first_string(value);
        if (genericStyleValue) {
          result.fontStyle = genericStyleValue;
        }
      }
      if (subcreator_visual_is_generic_font_style_key(normalizedKey)) {
        var genericStyleOptions = subcreator_visual_extract_string_options(value);
        for (var genericStyleIndex = 0; genericStyleIndex < genericStyleOptions.length; genericStyleIndex += 1) {
          subcreator_visual_push_unique_option(result.fontStyleOptions, genericStyleOptions[genericStyleIndex]);
          if (result.fontFamily) {
            subcreator_visual_register_style_for_family(result.fontStylesByFamily, result.fontFamily, genericStyleOptions[genericStyleIndex]);
          }
        }
      }

      if (isNaN(result.fontSize) && subcreator_visual_is_font_size_key(normalizedKey)) {
        var sizeValue = subcreator_visual_extract_first_number(value);
        if (!isNaN(sizeValue) && sizeValue > 0 && sizeValue < 2000) {
          result.fontSize = sizeValue;
        }
      }
      if (isNaN(result.fontSize) && subcreator_visual_is_generic_font_size_key(normalizedKey)) {
        var genericSizeValue = subcreator_visual_extract_first_number(value);
        if (!isNaN(genericSizeValue) && genericSizeValue > 0 && genericSizeValue < 2000) {
          result.fontSize = genericSizeValue;
        }
      }

      if (result.fontFauxStyleEditable === null && normalizedKey === "cappropfontfauxstyleedit") {
        // // Respect the text payload editability flags so the Visual Editor only exposes faux-style toggles when AE actually exported them.
        result.fontFauxStyleEditable = subcreator_visual_extract_first_boolean(value);
      }

      if (result.fontFsBold === null && subcreator_visual_is_font_flag_key(normalizedKey, "fontfsbold")) {
        result.fontFsBold = subcreator_visual_extract_first_boolean(value);
      }
      if (result.fontFsItalic === null && subcreator_visual_is_font_flag_key(normalizedKey, "fontfsitalic")) {
        result.fontFsItalic = subcreator_visual_extract_first_boolean(value);
      }
      if (result.fontFsAllCaps === null && subcreator_visual_is_font_flag_key(normalizedKey, "fontfsallcaps")) {
        result.fontFsAllCaps = subcreator_visual_extract_first_boolean(value);
      }
      if (result.fontFsSmallCaps === null && subcreator_visual_is_font_flag_key(normalizedKey, "fontfssmallcaps")) {
        result.fontFsSmallCaps = subcreator_visual_extract_first_boolean(value);
      }

      if (value && typeof value === "object") {
        scanNode(value, depth + 1);
      }
    }
  }

  scanNode(payload, 0);

  if (result.fontFamily) {
    subcreator_visual_push_unique_option(result.fontFamilyOptions, result.fontFamily);
  }
  if (result.fontStyle) {
    subcreator_visual_push_unique_option(result.fontStyleOptions, result.fontStyle);
  }
  if (result.fontFamily && result.fontStyle) {
    subcreator_visual_register_style_for_family(result.fontStylesByFamily, result.fontFamily, result.fontStyle);
  }

  if (result.fontFauxStyleEditable === false) {
    // // Hide Bold/Italic/All Caps/Small Caps when the MOGRT payload marks faux styles as non-editable.
    result.fontFsBold = null;
    result.fontFsItalic = null;
    result.fontFsAllCaps = null;
    result.fontFsSmallCaps = null;
  }

  if (!result.fontFamily && result.fontFamilyOptions.length === 1) {
    // // Keep font controls visible when readback returns only one family candidate.
    result.fontFamily = result.fontFamilyOptions[0];
  }
  if (!result.fontStyle && result.fontFamily && result.fontStylesByFamily[result.fontFamily] && result.fontStylesByFamily[result.fontFamily].length === 1) {
    // // Promote one unambiguous family style option to the current style value.
    result.fontStyle = result.fontStylesByFamily[result.fontFamily][0];
  }
  if (!result.fontStyle && result.fontStyleOptions.length === 1) {
    // // Keep one-style controls stable even when host omits the explicit current style.
    result.fontStyle = result.fontStyleOptions[0];
  }

  result.fontStylesByFamily = subcreator_visual_build_style_map_output(result.fontStylesByFamily);

  if (
    !result.fontFamily &&
    !result.fontStyle &&
    isNaN(result.fontSize) &&
    result.fontFamilyOptions.length < 1 &&
    result.fontStyleOptions.length < 1 &&
    !subcreator_visual_has_own_entries(result.fontStylesByFamily) &&
    result.fontFauxStyleEditable === null &&
    result.fontFsBold === null &&
    result.fontFsItalic === null &&
    result.fontFsAllCaps === null &&
    result.fontFsSmallCaps === null
  ) {
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

  function buildStringSelectOptions(primaryOptions, cacheOptions, fallbackOptions) {
    // // Merge primary/cache/fallback string options into panel select option descriptors.
    var merged = [];
    var sourceLists = [primaryOptions || [], cacheOptions || [], fallbackOptions || []];
    for (var listIndex = 0; listIndex < sourceLists.length; listIndex += 1) {
      var currentList = sourceLists[listIndex];
      for (var itemIndex = 0; itemIndex < currentList.length; itemIndex += 1) {
        subcreator_visual_push_unique_option(merged, currentList[itemIndex]);
      }
    }

    var descriptors = [];
    for (var optionIndex = 0; optionIndex < merged.length; optionIndex += 1) {
      descriptors.push({
        value: merged[optionIndex],
        label: merged[optionIndex]
      });
    }
    return descriptors;
  }

  if (styleValues.fontFamily) {
    subcreator_visual_register_text_style_option("family", styleValues.fontFamily);
  }
  if (styleValues.fontStyle) {
    subcreator_visual_register_text_style_option("style", styleValues.fontStyle);
  }
  for (var styleFamilyIndex = 0; styleFamilyIndex < styleValues.fontFamilyOptions.length; styleFamilyIndex += 1) {
    subcreator_visual_register_text_style_option("family", styleValues.fontFamilyOptions[styleFamilyIndex]);
  }
  for (var styleNameIndex = 0; styleNameIndex < styleValues.fontStyleOptions.length; styleNameIndex += 1) {
    subcreator_visual_register_text_style_option("style", styleValues.fontStyleOptions[styleNameIndex]);
  }

  var cachedFamilies = subcreator_visual_read_text_style_option_cache("family");
  var cachedStyles = subcreator_visual_read_text_style_option_cache("style");
  var commonStyleFallback = ["Regular", "Medium", "Semibold", "Bold", "Italic", "Bold Italic", "Black", "ExtraBold"];

  var entries = [];
  var targetGroup = groupPath || "General";

  if (styleValues.fontFamily || styleValues.fontFamilyOptions.length > 0) {
    var familyValue = styleValues.fontFamily || (styleValues.fontFamilyOptions.length > 0 ? styleValues.fontFamilyOptions[0] : "");
    entries.push({
      path: currentPath + "::textstyle.fontFamily",
      displayName: "Font Family",
      groupPath: targetGroup,
      valueType: "string",
      controlKind: "select",
      fontToken: styleValues.fontToken || "",
      options: buildStringSelectOptions(styleValues.fontFamilyOptions, cachedFamilies, []),
      styleOptionsByFamily: styleValues.fontStylesByFamily,
      value: familyValue
    });
  }

  if (
    styleValues.fontStyle ||
    styleValues.fontStyleOptions.length > 0 ||
    (styleValues.fontFamily &&
      styleValues.fontStylesByFamily[styleValues.fontFamily] &&
      styleValues.fontStylesByFamily[styleValues.fontFamily].length > 0)
  ) {
    var primaryFamily = styleValues.fontFamily || (styleValues.fontFamilyOptions.length > 0 ? styleValues.fontFamilyOptions[0] : "");
    var mappedStyleOptions =
      primaryFamily && styleValues.fontStylesByFamily[primaryFamily] && styleValues.fontStylesByFamily[primaryFamily].length > 0
        ? styleValues.fontStylesByFamily[primaryFamily]
        : styleValues.fontStyleOptions;
    var styleValue = styleValues.fontStyle || (mappedStyleOptions && mappedStyleOptions.length > 0 ? mappedStyleOptions[0] : "");
    entries.push({
      path: currentPath + "::textstyle.fontStyle",
      displayName: "Font Style",
      groupPath: targetGroup,
      valueType: "string",
      controlKind: "select",
      fontToken: styleValues.fontToken || "",
      options: buildStringSelectOptions(styleValues.fontStyleOptions, cachedStyles, commonStyleFallback),
      styleOptionsByFamily: styleValues.fontStylesByFamily,
      value: styleValue
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

  if (typeof styleValues.fontFsBold === "boolean") {
    entries.push({
      path: currentPath + "::textstyle.fontFsBold",
      displayName: "Bold",
      groupPath: targetGroup,
      valueType: "boolean",
      controlKind: "checkbox",
      value: styleValues.fontFsBold
    });
  }

  if (typeof styleValues.fontFsItalic === "boolean") {
    entries.push({
      path: currentPath + "::textstyle.fontFsItalic",
      displayName: "Italic",
      groupPath: targetGroup,
      valueType: "boolean",
      controlKind: "checkbox",
      value: styleValues.fontFsItalic
    });
  }

  if (typeof styleValues.fontFsAllCaps === "boolean") {
    entries.push({
      path: currentPath + "::textstyle.fontFsAllCaps",
      displayName: "All Caps",
      groupPath: targetGroup,
      valueType: "boolean",
      controlKind: "checkbox",
      value: styleValues.fontFsAllCaps
    });
  }

  if (typeof styleValues.fontFsSmallCaps === "boolean") {
    entries.push({
      path: currentPath + "::textstyle.fontFsSmallCaps",
      displayName: "Small Caps",
      groupPath: targetGroup,
      valueType: "boolean",
      controlKind: "checkbox",
      value: styleValues.fontFsSmallCaps
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

    var selectOptions = subcreator_visual_build_select_options(displayName, value, groupPath);
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
    var displayVectorRaw = subcreator_visual_vector_to_panel_units(vectorValue, vectorMeta.scale);
    var displayVector = [];
    for (var vectorDisplayIndex = 0; vectorDisplayIndex < displayVectorRaw.length; vectorDisplayIndex += 1) {
      displayVector.push(subcreator_visual_round_number_for_display(displayVectorRaw[vectorDisplayIndex], 1));
    }
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
        panelRaw: displayVectorRaw,
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

    var currentPath = subcreator_visual_join_property_path(pathPrefix, index);
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

function subcreator_visual_preview_debug_value(rawValue, maxLength) {
  // // Build compact readable previews for debug logs.
  var limit = Math.max(Number(maxLength || 280), 40);
  var preview = "";
  try {
    if (typeof rawValue === "string") {
      preview = rawValue;
    } else {
      preview = JSON.stringify(rawValue);
    }
  } catch (error) {
    preview = String(rawValue);
  }

  preview = String(preview || "").replace(/\r/g, "\\r").replace(/\n/g, "\\n");
  if (preview.length > limit) {
    preview = preview.substring(0, limit) + "...";
  }
  return preview;
}

function subcreator_collect_text_style_debug_candidates(propertyCollection, pathPrefix, groupPathPrefix, outList, maxItems) {
  // // Gather text-style detection hints (including skipped properties) to debug template-specific payloads.
  if (!propertyCollection || typeof propertyCollection.numItems !== "number" || !outList) {
    return;
  }

  var activeSiblingGroupPath = String(groupPathPrefix || "");
  var limit = Math.max(Number(maxItems || 20), 1);

  for (var index = 0; index < propertyCollection.numItems; index += 1) {
    if (outList.length >= limit) {
      return;
    }

    var property = propertyCollection[index];
    if (!property) {
      continue;
    }

    var currentPath = subcreator_visual_join_property_path(pathPrefix, index);
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
      } catch (readError) {
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

    if (hasValue && subcreator_visual_is_guid_list_string(rawValue)) {
      activeSiblingGroupPath = groupPathPrefix ? groupPathPrefix + " / " + displayName : displayName;
      if (hasChildren) {
        subcreator_collect_text_style_debug_candidates(
          property.properties,
          currentPath,
          activeSiblingGroupPath,
          outList,
          limit
        );
      }
      continue;
    }

    var resolvedGroupPath = activeSiblingGroupPath || groupPathPrefix || "General";
    if (hasValue) {
      var styleValues = subcreator_visual_extract_text_style_from_value(rawValue);
      var textCandidate = subcreator_should_try_text_property(displayName, rawValue);
      var nameKey = subcreator_visual_normalize_label_key(displayName);
      var maybeTextByLabel =
        nameKey.indexOf("text") !== -1 ||
        nameKey.indexOf("title") !== -1 ||
        nameKey.indexOf("name") !== -1 ||
        nameKey.indexOf("source") !== -1;

      if (styleValues || textCandidate || maybeTextByLabel) {
        outList.push({
          path: currentPath,
          name: displayName,
          group: resolvedGroupPath,
          hasChildren: hasChildren,
          hasSetter: typeof property.setValue === "function",
          rawType: typeof rawValue,
          rawPreview: subcreator_visual_preview_debug_value(rawValue, 260),
          styleDetected: styleValues || null
        });
      }
    }

    if (hasChildren) {
      var nextGroupPath = resolvedGroupPath ? resolvedGroupPath + " / " + displayName : displayName;
      subcreator_collect_text_style_debug_candidates(property.properties, currentPath, nextGroupPath, outList, limit);
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

function subcreator_visual_join_property_path(pathPrefix, index) {
  // // Join nested property indexes while keeping component-prefixed roots like `c1|0.2` stable.
  var prefix = subcreator_trim_string(String(pathPrefix || ""));
  if (!prefix) {
    return String(index);
  }

  if (prefix.charAt(prefix.length - 1) === "|") {
    return prefix + String(index);
  }

  return prefix + "." + String(index);
}

function subcreator_visual_build_component_path_prefix(componentIndex) {
  // // Prefix visual-editor paths with the source component index so Premiere-authored MOGRTs can expose multiple components safely.
  return "c" + String(Math.max(0, Number(componentIndex || 0))) + "|";
}

function subcreator_visual_get_component_group_label(component, componentIndex) {
  // // Use the component/layer name as the top-level visual-editor group label when multiple components are exposed.
  var componentName = subcreator_trim_string(String((component && (component.displayName || component.name)) || ""));
  if (!componentName) {
    return "Component " + String(Math.max(0, Number(componentIndex || 0)) + 1);
  }

  return componentName;
}

function subcreator_visual_group_mentions(groupPath, token) {
  // // Match simple component/group keywords without relying on exact localized group labels.
  var normalizedGroup = " " + String(groupPath || "").toLowerCase().replace(/[\/|_-]+/g, " ") + " ";
  var normalizedToken = subcreator_trim_string(String(token || "")).toLowerCase();
  if (!normalizedToken) {
    return false;
  }

  return normalizedGroup.indexOf(" " + normalizedToken + " ") !== -1;
}

function subcreator_visual_append_group_suffix(groupPath, suffix) {
  // // Preserve the top-level component label while splitting Premiere-only sections into clearer sub-groups.
  var baseGroup = subcreator_trim_string(String(groupPath || ""));
  var groupSuffix = subcreator_trim_string(String(suffix || ""));
  if (!groupSuffix) {
    return baseGroup || "General";
  }
  if (!baseGroup) {
    return groupSuffix;
  }
  if (String(baseGroup).toLowerCase().indexOf(String(groupSuffix).toLowerCase()) !== -1) {
    return baseGroup;
  }
  return baseGroup + " / " + groupSuffix;
}

function subcreator_visual_normalize_descriptor(descriptor) {
  // // Rename a few Premiere-authored internal labels into something closer to the Properties panel wording.
  if (!descriptor) {
    return null;
  }

  var normalized = {};
  for (var key in descriptor) {
    if (descriptor.hasOwnProperty(key)) {
      normalized[key] = descriptor[key];
    }
  }

  var displayName = subcreator_trim_string(String(normalized.displayName || ""));
  var groupPath = subcreator_trim_string(String(normalized.groupPath || ""));
  var normalizedKey = subcreator_visual_normalize_label_key(displayName);
  var isShapeGroup = subcreator_visual_group_mentions(groupPath, "shape");
  var isTextGroup = subcreator_visual_group_mentions(groupPath, "text");

  if ((isShapeGroup || isTextGroup) && /^(left|top|right|bottom)$/.test(normalizedKey)) {
    normalized.groupPath = subcreator_visual_append_group_suffix(groupPath, "Responsive Design");
  }

  return normalized;
}

function subcreator_visual_should_hide_descriptor(descriptor) {
  // // Hide internal or misleading controls so Premiere-authored templates stay readable in the Visual editor.
  if (!descriptor) {
    return true;
  }

  var displayName = subcreator_trim_string(String(descriptor.displayName || ""));
  var groupPath = String(descriptor.groupPath || "");
  var normalizedKey = subcreator_visual_normalize_label_key(displayName);
  if (!displayName || /^Property\s+/i.test(displayName)) {
    return true;
  }

  if (subcreator_visual_group_mentions(groupPath, "responsive")) {
    return true;
  }

  if (normalizedKey === "align") {
    return true;
  }

  if (normalizedKey === "roundedcrop") {
    // // Premiere can expose this internal crop helper only after some write/refresh cycles even when the MOGRT never exposed it in Essential Graphics.
    return true;
  }

  if (
    normalizedKey === "parentwidth" ||
    normalizedKey === "parentheight" ||
    normalizedKey === "parentrotation"
  ) {
    return true;
  }

  if ((normalizedKey === "start" || normalizedKey === "end") && subcreator_visual_group_mentions(groupPath, "text")) {
    return true;
  }

  if (/^(left|top|right|bottom)$/.test(normalizedKey) && (subcreator_visual_group_mentions(groupPath, "shape") || subcreator_visual_group_mentions(groupPath, "text"))) {
    return true;
  }

  if (
    (normalizedKey === "path" || normalizedKey === "appearance" || normalizedKey === "transform") &&
    descriptor.controlKind === "string"
  ) {
    return true;
  }

  if (normalizedKey === "transform" && descriptor.controlKind === "boolean") {
    return true;
  }

  if (
    normalizedKey === "erroroccurred" ||
    normalizedKey === "controls" ||
    normalizedKey === "appliedversion" ||
    normalizedKey === "sequencewidth" ||
    normalizedKey === "sequenceheight" ||
    normalizedKey === "sequencepixelratio"
  ) {
    return true;
  }

  if (normalizedKey === "blendmode" && subcreator_visual_group_mentions(groupPath, "opacity")) {
    // // Hide clip-level blend mode until CEP exposes a reliable read/write mapping; current values drift from Premiere's visible label and writes do not stick.
    return true;
  }

  return false;
}

function subcreator_visual_filter_property_descriptors(properties) {
  // // Remove duplicate or low-signal descriptors from the Visual editor list while keeping host paths stable.
  if (!properties || typeof properties.length !== "number") {
    return [];
  }

  var filtered = [];
  var seenGroupDisplayKeys = {};

  for (var index = 0; index < properties.length; index += 1) {
    var descriptor = subcreator_visual_normalize_descriptor(properties[index]);
    if (subcreator_visual_should_hide_descriptor(descriptor)) {
      continue;
    }

    var displayName = subcreator_trim_string(String(descriptor.displayName || ""));
    var groupPath = subcreator_trim_string(String(descriptor.groupPath || ""));
    var duplicateKey = String(groupPath).toLowerCase() + "::" + String(displayName).toLowerCase();
    if (String(displayName).toLowerCase() === "blend mode") {
      if (seenGroupDisplayKeys[duplicateKey]) {
        continue;
      }
      seenGroupDisplayKeys[duplicateKey] = true;
    }

    filtered.push(descriptor);
  }

  return filtered;
}

function subcreator_visual_build_component_descriptor_signature(properties) {
  // // Deduplicate cloned AE components by visible descriptor content instead of host path prefix.
  if (!properties || typeof properties.length !== "number") {
    return "";
  }

  var parts = [];
  for (var index = 0; index < properties.length; index += 1) {
    var descriptor = properties[index];
    if (!descriptor) {
      continue;
    }

    var normalizedDescriptor = subcreator_visual_normalize_descriptor(descriptor) || descriptor;
    var optionParts = [];
    if (normalizedDescriptor.options && typeof normalizedDescriptor.options.length === "number") {
      for (var optionIndex = 0; optionIndex < normalizedDescriptor.options.length; optionIndex += 1) {
        var option = normalizedDescriptor.options[optionIndex] || {};
        optionParts.push(String(option.value) + "=" + String(option.label || option.value));
      }
    }

    var vectorScaleText = "";
    if (normalizedDescriptor.vectorScale && typeof normalizedDescriptor.vectorScale.length === "number") {
      vectorScaleText = String(normalizedDescriptor.vectorScale.join(","));
    }

    parts.push(
      [
        String(normalizedDescriptor.displayName || ""),
        String(normalizedDescriptor.groupPath || ""),
        String(normalizedDescriptor.controlKind || ""),
        String(normalizedDescriptor.valueType || ""),
        String(normalizedDescriptor.value),
        String(normalizedDescriptor.minValue),
        String(normalizedDescriptor.maxValue),
        String(normalizedDescriptor.stepValue),
        String(normalizedDescriptor.vectorMode || ""),
        vectorScaleText,
        optionParts.join("|")
      ].join("::")
    );
  }

  return parts.join("\n");
}

function subcreator_collect_unique_mogrt_components(trackItem) {
  // // Reuse the same visible-descriptor deduplication as the Visual editor so hidden duplicate AE components do not receive conflicting writes.
  var rawComponents = subcreator_get_mogrt_components_from_track_item(trackItem);
  if (rawComponents.length < 2) {
    return rawComponents;
  }

  var uniqueComponents = [];
  var seenComponentSignatures = {};
  subcreator_visual_reset_group_sequence_axis_preferences();
  subcreator_visual_reset_text_style_option_cache();

  for (var componentIndex = 0; componentIndex < rawComponents.length; componentIndex += 1) {
    var rawComponent = rawComponents[componentIndex];
    if (!rawComponent || !rawComponent.properties || rawComponent.properties.numItems < 1) {
      continue;
    }

    var componentGroupPath = rawComponents.length > 1 ? subcreator_visual_get_component_group_label(rawComponent, componentIndex) : "";
    var componentProperties = [];
    subcreator_collect_mogrt_visual_properties_recursive(
      rawComponent.properties,
      subcreator_visual_build_component_path_prefix(componentIndex),
      componentGroupPath,
      componentProperties
    );

    componentProperties = subcreator_visual_filter_property_descriptors(componentProperties);
    var componentSignature = subcreator_visual_build_component_descriptor_signature(componentProperties);
    if (componentSignature && typeof seenComponentSignatures[componentSignature] !== "undefined") {
      continue;
    }

    seenComponentSignatures[componentSignature] = componentIndex;
    uniqueComponents.push(rawComponent);
  }

  return uniqueComponents;
}

function subcreator_visual_parse_component_prefixed_path(pathValue) {
  // // Decode optional component-prefixed visual-editor paths like `c2|0.4::textstyle.fontSize`.
  var pathText = subcreator_trim_string(String(pathValue || ""));
  var match = /^c(\d+)\|([\s\S]*)$/.exec(pathText);
  if (!match) {
    return {
      componentIndex: 0,
      propertyPath: pathText
    };
  }

  return {
    componentIndex: Number(match[1] || 0),
    propertyPath: subcreator_trim_string(String(match[2] || ""))
  };
}

function subcreator_visual_resolve_property_from_track_item(trackItem, pathValue) {
  // // Resolve one visual-editor path back to the correct component/property pair on a track item.
  var parsedPath = subcreator_visual_parse_component_prefixed_path(pathValue);
  var components = subcreator_get_mogrt_components_from_track_item(trackItem);
  var component = components[parsedPath.componentIndex];
  if (!component || !component.properties) {
    return null;
  }

  return {
    componentIndex: parsedPath.componentIndex,
    propertyPath: parsedPath.propertyPath,
    component: component,
    property: subcreator_find_property_by_path(component.properties, parsedPath.propertyPath)
  };
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

  if (
    normalizedStyleKey === "fontFsBold" ||
    normalizedStyleKey === "fontFsItalic" ||
    normalizedStyleKey === "fontFsAllCaps" ||
    normalizedStyleKey === "fontFsSmallCaps"
  ) {
    if (typeof value === "boolean") {
      return value;
    }
    var boolText = subcreator_trim_string(String(value || "")).toLowerCase();
    if (boolText === "true" || boolText === "1" || boolText === "yes") {
      return true;
    }
    if (boolText === "false" || boolText === "0" || boolText === "no") {
      return false;
    }
    return null;
  }

  if (normalizedStyleKey === "fontSize") {
    var sizeValue = Number(value);
    if (isNaN(sizeValue) || sizeValue <= 0 || sizeValue > 2000) {
      return null;
    }
    return sizeValue;
  }

  if (normalizedStyleKey === "fontStyle") {
    // // Allow an explicit empty style so cloned source tokens like `Impact` can clear a target's leftover `Bold` style.
    if (value === "") {
      return "";
    }
    if (typeof value === "string" && subcreator_trim_string(value) === "") {
      return "";
    }
  }

  var textValue = subcreator_trim_string(String(value || ""));
  if (!textValue) {
    return null;
  }
  return textValue;
}

function subcreator_visual_apply_text_style_to_payload(payload, styleKey, styleValue, applyOptions) {
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

      if (styleKey === "fontFsBold" && subcreator_visual_is_font_flag_key(normalizedKey, "fontfsbold")) {
        if (typeof value === "boolean") {
          node[key] = styleValue === true;
        } else if (value && typeof value.length === "number" && value.length > 0) {
          value[0] = styleValue === true;
        } else {
          node[key] = [styleValue === true];
        }
        updated = true;
      } else if (styleKey === "fontFsItalic" && subcreator_visual_is_font_flag_key(normalizedKey, "fontfsitalic")) {
        if (typeof value === "boolean") {
          node[key] = styleValue === true;
        } else if (value && typeof value.length === "number" && value.length > 0) {
          value[0] = styleValue === true;
        } else {
          node[key] = [styleValue === true];
        }
        updated = true;
      } else if (styleKey === "fontFsAllCaps" && subcreator_visual_is_font_flag_key(normalizedKey, "fontfsallcaps")) {
        if (typeof value === "boolean") {
          node[key] = styleValue === true;
        } else if (value && typeof value.length === "number" && value.length > 0) {
          value[0] = styleValue === true;
        } else {
          node[key] = [styleValue === true];
        }
        updated = true;
      } else if (styleKey === "fontFsSmallCaps" && subcreator_visual_is_font_flag_key(normalizedKey, "fontfssmallcaps")) {
        if (typeof value === "boolean") {
          node[key] = styleValue === true;
        } else if (value && typeof value.length === "number" && value.length > 0) {
          value[0] = styleValue === true;
        } else {
          node[key] = [styleValue === true];
        }
        updated = true;
      } else if (styleKey === "fontFamily" && subcreator_visual_is_generic_font_family_key(normalizedKey)) {
        var currentToken = subcreator_visual_extract_first_string(value);
        var currentParts = subcreator_visual_split_font_token(currentToken);
        var tokenOverrideProvided =
          applyOptions &&
          Object.prototype.hasOwnProperty.call(applyOptions, "fontTokenOverride") &&
          typeof applyOptions.fontTokenOverride === "string";
        var styleOverrideProvided =
          applyOptions &&
          Object.prototype.hasOwnProperty.call(applyOptions, "fontTokenStyleOverride") &&
          typeof applyOptions.fontTokenStyleOverride === "string";
        var styleForToken = styleOverrideProvided ? applyOptions.fontTokenStyleOverride : currentParts.style;
        var rebuiltToken = tokenOverrideProvided
          ? applyOptions.fontTokenOverride
          : subcreator_visual_join_font_token(styleValue, styleForToken, currentToken);
        if (typeof value === "string") {
          node[key] = rebuiltToken;
        } else if (value && typeof value.length === "number" && value.length > 0) {
          value[0] = rebuiltToken;
        } else {
          node[key] = [rebuiltToken];
        }
        updated = true;
      } else if (styleKey === "fontStyle" && subcreator_visual_is_generic_font_family_key(normalizedKey)) {
        var existingToken = subcreator_visual_extract_first_string(value);
        var existingParts = subcreator_visual_split_font_token(existingToken);
        var tokenOverrideForStyleProvided =
          applyOptions &&
          Object.prototype.hasOwnProperty.call(applyOptions, "fontTokenOverride") &&
          typeof applyOptions.fontTokenOverride === "string";
        var familyOverrideProvided =
          applyOptions &&
          Object.prototype.hasOwnProperty.call(applyOptions, "fontTokenFamilyOverride") &&
          typeof applyOptions.fontTokenFamilyOverride === "string";
        var familyForToken = familyOverrideProvided ? applyOptions.fontTokenFamilyOverride : existingParts.family;
        var rebuiltStyleToken = tokenOverrideForStyleProvided
          ? applyOptions.fontTokenOverride
          : subcreator_visual_join_font_token(familyForToken, styleValue, existingToken);
        if (typeof value === "string") {
          node[key] = rebuiltStyleToken;
        } else if (value && typeof value.length === "number" && value.length > 0) {
          value[0] = rebuiltStyleToken;
        } else {
          node[key] = [rebuiltStyleToken];
        }
        updated = true;
      } else if (styleKey === "fontFamily" && subcreator_visual_is_font_family_key(normalizedKey)) {
        node[key] = String(styleValue);
        updated = true;
      } else if (styleKey === "fontStyle" && subcreator_visual_is_font_style_key(normalizedKey)) {
        node[key] = String(styleValue);
        updated = true;
      } else if (styleKey === "fontStyle" && subcreator_visual_is_generic_font_style_key(normalizedKey)) {
        if (typeof value === "string") {
          node[key] = String(styleValue);
        } else if (value && typeof value.length === "number" && value.length > 0) {
          value[0] = String(styleValue);
        } else {
          node[key] = [String(styleValue)];
        }
        updated = true;
      } else if (styleKey === "fontSize" && subcreator_visual_is_font_size_key(normalizedKey)) {
        if (typeof value === "string" && value !== "") {
          node[key] = String(Number(styleValue));
        } else if (value && typeof value.length === "number" && value.length > 0) {
          value[0] = Number(styleValue);
        } else {
          node[key] = Number(styleValue);
        }
        updated = true;
      } else if (styleKey === "fontSize" && subcreator_visual_is_generic_font_size_key(normalizedKey)) {
        if (typeof value === "string" && value !== "") {
          node[key] = String(Number(styleValue));
        } else if (value && typeof value.length === "number" && value.length > 0) {
          value[0] = Number(styleValue);
        } else {
          node[key] = Number(styleValue);
        }
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

function subcreator_try_patch_text_style_json_string(rawValue, styleKey, styleValue, applyOptions) {
  // // Fallback patch for JSON-like strings when `JSON.parse` is unavailable on host payloads.
  var raw = String(rawValue || "");
  if (!raw || raw.indexOf("{") === -1) {
    return "";
  }

  var patched = raw;
  var styleString = JSON.stringify(String(styleValue));
  var tokenOverrideProvided =
    applyOptions &&
    Object.prototype.hasOwnProperty.call(applyOptions, "fontTokenOverride") &&
    typeof applyOptions.fontTokenOverride === "string";
  var tokenOverrideValue = tokenOverrideProvided ? applyOptions.fontTokenOverride : "";
  var styleOverrideProvided =
    applyOptions &&
    Object.prototype.hasOwnProperty.call(applyOptions, "fontTokenStyleOverride") &&
    typeof applyOptions.fontTokenStyleOverride === "string";
  var styleOverrideValue = styleOverrideProvided ? applyOptions.fontTokenStyleOverride : "";
  var familyOverrideProvided =
    applyOptions &&
    Object.prototype.hasOwnProperty.call(applyOptions, "fontTokenFamilyOverride") &&
    typeof applyOptions.fontTokenFamilyOverride === "string";
  var familyOverrideValue = familyOverrideProvided ? applyOptions.fontTokenFamilyOverride : "";
  var keyList = [];

  if (styleKey === "fontFamily") {
    keyList = ["fontName", "mFontName", "fontFamily", "mFontFamily", "fontEditValue", "mFontEditValue"];
  } else if (styleKey === "fontStyle") {
    keyList = [
      "fontStyle",
      "mFontStyle",
      "fontStyleName",
      "mFontStyleName",
      "fontStyleValue",
      "fontStyleEditValue",
      "fontFauxStyleValue",
      "fontFauxStyleEditValue"
    ];
  } else if (styleKey === "fontSize") {
    keyList = ["fontSize", "mFontSize", "fontSizeValue", "fontSizeEditValue", "mFontSizeValue", "mFontSizeEditValue"];
  } else if (styleKey === "fontFsBold") {
    keyList = ["fontFSBoldValue", "fontFSBoldEditValue", "mFontFSBoldValue", "mFontFSBoldEditValue"];
  } else if (styleKey === "fontFsItalic") {
    keyList = ["fontFSItalicValue", "fontFSItalicEditValue", "mFontFSItalicValue", "mFontFSItalicEditValue"];
  } else if (styleKey === "fontFsAllCaps") {
    keyList = ["fontFSAllCapsValue", "fontFSAllCapsEditValue", "mFontFSAllCapsValue", "mFontFSAllCapsEditValue"];
  } else if (styleKey === "fontFsSmallCaps") {
    keyList = ["fontFSSmallCapsValue", "fontFSSmallCapsEditValue", "mFontFSSmallCapsValue", "mFontFSSmallCapsEditValue"];
  }

  for (var keyIndex = 0; keyIndex < keyList.length; keyIndex += 1) {
    var keyName = keyList[keyIndex];
    if (styleKey === "fontSize") {
      var numericRegex = new RegExp('"' + keyName + '"\\s*:\\s*("([^"\\\\]|\\\\.)*"|-?\\d+(?:\\.\\d+)?)', "g");
      patched = patched.replace(numericRegex, '"' + keyName + '":' + String(Number(styleValue)));
      var numericArrayRegex = new RegExp('"' + keyName + '"\\s*:\\s*\\[[^\\]]*\\]', "g");
      patched = patched.replace(numericArrayRegex, '"' + keyName + '":[' + String(Number(styleValue)) + "]");
    } else {
      if (keyName === "fontEditValue" || keyName === "mFontEditValue") {
        var fontTokenRegex = new RegExp('"' + keyName + '"\\s*:\\s*\\["((?:[^"\\\\]|\\\\.)*)"\\]', "g");
        patched = patched.replace(fontTokenRegex, function (matchValue, tokenValue) {
          var tokenParts = subcreator_visual_split_font_token(String(tokenValue || ""));
          var rebuiltToken = "";
          if (tokenOverrideProvided) {
            rebuiltToken = tokenOverrideValue;
          } else if (styleKey === "fontFamily") {
            var styleForToken = styleOverrideProvided ? styleOverrideValue : tokenParts.style;
            rebuiltToken = subcreator_visual_join_font_token(String(styleValue), styleForToken, tokenValue);
          } else if (styleKey === "fontStyle") {
            var familyForToken = familyOverrideProvided ? familyOverrideValue : tokenParts.family;
            rebuiltToken = subcreator_visual_join_font_token(familyForToken, String(styleValue), tokenValue);
          } else {
            rebuiltToken = String(styleValue);
          }
          return '"' + keyName + '":[' + JSON.stringify(rebuiltToken) + "]";
        });
        continue;
      }
      if (
        styleKey === "fontFsBold" ||
        styleKey === "fontFsItalic" ||
        styleKey === "fontFsAllCaps" ||
        styleKey === "fontFsSmallCaps"
      ) {
        var boolValue = styleValue === true ? "true" : "false";
        var boolRegex = new RegExp('"' + keyName + '"\\s*:\\s*(true|false|0|1)', "g");
        patched = patched.replace(boolRegex, '"' + keyName + '":' + boolValue);
        var boolArrayRegex = new RegExp('"' + keyName + '"\\s*:\\s*\\[[^\\]]*\\]', "g");
        patched = patched.replace(boolArrayRegex, '"' + keyName + '":[' + boolValue + "]");
        continue;
      }
      var stringRegex = new RegExp('"' + keyName + '"\\s*:\\s*"([^"\\\\]|\\\\.)*"', "g");
      patched = patched.replace(stringRegex, '"' + keyName + '":' + styleString);
      var stringArrayRegex = new RegExp('"' + keyName + '"\\s*:\\s*\\["([^"\\\\]|\\\\.)*"\\]', "g");
      patched = patched.replace(stringArrayRegex, '"' + keyName + '":[' + styleString + "]");
    }
  }

  if (patched === raw) {
    return "";
  }

  return patched;
}

function subcreator_try_set_mogrt_text_style_property(property, styleKey, styleValue, extraOptions) {
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

  var exclusiveCompanionStyleKey = "";
  // // Keep faux-style exclusive toggles aligned with Premiere behavior.
  if (normalizedStyleKey === "fontFsAllCaps" && normalizedStyleValue === true) {
    exclusiveCompanionStyleKey = "fontFsSmallCaps";
  } else if (normalizedStyleKey === "fontFsSmallCaps" && normalizedStyleValue === true) {
    exclusiveCompanionStyleKey = "fontFsAllCaps";
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

  var extractedStyleValues = subcreator_visual_extract_text_style_from_value(rawValue);

  function applyStylePatch(styleKeyToApply, baseStyleValue, customOptions) {
    // // Try one text-style patch strategy for any text-style key, including companion font-style cleanup after a family clone.
    var effectiveStyleKey = subcreator_trim_string(String(styleKeyToApply || normalizedStyleKey));
    if (!effectiveStyleKey) {
      return false;
    }

    var effectiveStyleValue = baseStyleValue;
    var applyOptions = null;
    if (customOptions && typeof customOptions === "object") {
      if (Object.prototype.hasOwnProperty.call(customOptions, "styleValueOverride")) {
        var overrideNormalized = subcreator_visual_normalize_text_style_change(
          effectiveStyleKey,
          customOptions.styleValueOverride
        );
        if (overrideNormalized === null) {
          return false;
        }
        effectiveStyleValue = overrideNormalized;
      }
      if (
        typeof customOptions.fontTokenOverride === "string" ||
        typeof customOptions.fontTokenStyleOverride === "string" ||
        typeof customOptions.fontTokenFamilyOverride === "string"
      ) {
        applyOptions = {};
        if (typeof customOptions.fontTokenStyleOverride === "string") {
          applyOptions.fontTokenStyleOverride = customOptions.fontTokenStyleOverride;
        }
        if (typeof customOptions.fontTokenFamilyOverride === "string") {
          applyOptions.fontTokenFamilyOverride = customOptions.fontTokenFamilyOverride;
        }
        if (typeof customOptions.fontTokenOverride === "string") {
          applyOptions.fontTokenOverride = customOptions.fontTokenOverride;
        }
      }
    }

    if (rawValue && typeof rawValue === "object") {
      try {
        var objectCopy = JSON.parse(JSON.stringify(rawValue));
        var didPatchCopy = subcreator_visual_apply_text_style_to_payload(
          objectCopy,
          effectiveStyleKey,
          effectiveStyleValue,
          applyOptions
        );
        if (didPatchCopy && exclusiveCompanionStyleKey && effectiveStyleKey === normalizedStyleKey) {
          subcreator_visual_apply_text_style_to_payload(objectCopy, exclusiveCompanionStyleKey, false, applyOptions);
        }
        if (didPatchCopy) {
          property.setValue(objectCopy, true);
          return true;
        }
      } catch (copyError) {}

      try {
        var didPatchDirect = subcreator_visual_apply_text_style_to_payload(
          rawValue,
          effectiveStyleKey,
          effectiveStyleValue,
          applyOptions
        );
        if (didPatchDirect && exclusiveCompanionStyleKey && effectiveStyleKey === normalizedStyleKey) {
          subcreator_visual_apply_text_style_to_payload(rawValue, exclusiveCompanionStyleKey, false, applyOptions);
        }
        if (didPatchDirect) {
          property.setValue(rawValue, true);
          return true;
        }
      } catch (directError) {}
    }

    if (typeof rawValue === "string" && rawValue.indexOf("{") !== -1) {
      try {
        var parsed = JSON.parse(rawValue);
        var didPatchParsed = subcreator_visual_apply_text_style_to_payload(
          parsed,
          effectiveStyleKey,
          effectiveStyleValue,
          applyOptions
        );
        if (didPatchParsed && exclusiveCompanionStyleKey && effectiveStyleKey === normalizedStyleKey) {
          subcreator_visual_apply_text_style_to_payload(parsed, exclusiveCompanionStyleKey, false, applyOptions);
        }
        if (didPatchParsed) {
          property.setValue(JSON.stringify(parsed), true);
          return true;
        }
      } catch (jsonError) {}

      try {
        var patchedRaw = subcreator_try_patch_text_style_json_string(
          rawValue,
          effectiveStyleKey,
          effectiveStyleValue,
          applyOptions
        );
        if (patchedRaw && exclusiveCompanionStyleKey && effectiveStyleKey === normalizedStyleKey) {
          patchedRaw = subcreator_try_patch_text_style_json_string(patchedRaw, exclusiveCompanionStyleKey, false, applyOptions) || patchedRaw;
        }
        if (patchedRaw) {
          property.setValue(patchedRaw, true);
          return true;
        }
      } catch (patchError) {}
    }

    return false;
  }

  function applyOnce(customOptions) {
    // // Reuse the generalized patch helper for the main requested style key.
    return applyStylePatch(normalizedStyleKey, normalizedStyleValue, customOptions);
  }

  function readbackMatchesExpectedFamily(expectedFamily) {
    // // Validate family write result; return `false` only when mismatch is explicit.
    if (typeof property.getValue !== "function") {
      return null;
    }
    try {
      var afterValue = property.getValue();
      return subcreator_visual_font_family_matches_value(afterValue, expectedFamily);
    } catch (readError) {
      return null;
    }
  }

  if (normalizedStyleKey === "fontFamily") {
    var requestedFamily = String(normalizedStyleValue || "");
    var requestedFamilyKey = subcreator_visual_normalize_font_compare_key(requestedFamily);
    var suppliedFontToken = extraOptions && typeof extraOptions.fontToken === "string" ? extraOptions.fontToken : "";
    var suppliedTokenParts = subcreator_visual_split_font_token(suppliedFontToken);
    var desiredCompanionStyle = suppliedTokenParts && suppliedTokenParts.style ? suppliedTokenParts.style : "";
    var styleOverrides = [];

    function pushStyleOverride(value) {
      // // Keep fallback style overrides unique while preserving insertion order.
      var text = subcreator_trim_string(String(value || ""));
      var normalized = text.toLowerCase();
      for (var idx = 0; idx < styleOverrides.length; idx += 1) {
        if (String(styleOverrides[idx] || "").toLowerCase() === normalized) {
          return;
        }
      }
      styleOverrides.push(text);
    }

    if (suppliedTokenParts && suppliedTokenParts.style) {
      pushStyleOverride(suppliedTokenParts.style);
    }

    if (
      suppliedFontToken &&
      suppliedTokenParts &&
      suppliedTokenParts.family &&
      subcreator_visual_normalize_font_compare_key(suppliedTokenParts.family) === requestedFamilyKey
    ) {
      // // Try the exact source token first so copy/apply keeps fonts like `Impact` instead of inventing `Impact Regular`.
      if (
        applyOnce({
          fontTokenOverride: suppliedFontToken
        })
      ) {
        // // Align any separate style fields with the source token so target clips do not keep a stale `Bold` style.
        applyStylePatch("fontStyle", desiredCompanionStyle, {
          styleValueOverride: desiredCompanionStyle,
          fontTokenFamilyOverride: requestedFamily,
          fontTokenOverride: suppliedFontToken
        });
        var directTokenReadbackMatch = readbackMatchesExpectedFamily(requestedFamily);
        if (directTokenReadbackMatch !== false) {
          return true;
        }
      }
    }

    // // Prefer neutral defaults for a newly selected family before reusing the previous preset style.
    pushStyleOverride("Regular");
    pushStyleOverride("Plain");
    pushStyleOverride("Roman");
    pushStyleOverride("Book");
    pushStyleOverride("Medium");

    if (extractedStyleValues && extractedStyleValues.fontStylesByFamily && requestedFamilyKey) {
      for (var familyKey in extractedStyleValues.fontStylesByFamily) {
        if (!extractedStyleValues.fontStylesByFamily.hasOwnProperty(familyKey)) {
          continue;
        }
        if (subcreator_visual_normalize_font_compare_key(familyKey) !== requestedFamilyKey) {
          continue;
        }
        var styleBucket = extractedStyleValues.fontStylesByFamily[familyKey];
        if (!styleBucket || typeof styleBucket.length !== "number") {
          continue;
        }
        for (var styleIndex = 0; styleIndex < styleBucket.length; styleIndex += 1) {
          pushStyleOverride(styleBucket[styleIndex]);
        }
      }
    }

    if (extractedStyleValues && extractedStyleValues.fontStyle) {
      pushStyleOverride(extractedStyleValues.fontStyle);
    }

    // // Try no-style last because some families need an explicit regular token to avoid visual fallback fonts.
    pushStyleOverride("");

    if (!styleOverrides.length) {
      pushStyleOverride("");
    }

    for (var overrideIndex = 0; overrideIndex < styleOverrides.length; overrideIndex += 1) {
      var styleOverride = styleOverrides[overrideIndex];
      var tokenCandidates = subcreator_visual_build_font_token_candidates(requestedFamily, styleOverride, suppliedFontToken, "");
      if (tokenCandidates.length < 1) {
        tokenCandidates.push("");
      }
      for (var tokenCandidateIndex = 0; tokenCandidateIndex < tokenCandidates.length; tokenCandidateIndex += 1) {
        var tokenCandidate = tokenCandidates[tokenCandidateIndex];
        var applyConfig = {
          fontTokenStyleOverride: styleOverride
        };
        if (tokenCandidate) {
          applyConfig.fontTokenOverride = tokenCandidate;
        }
        if (!applyOnce(applyConfig)) {
          continue;
        }
        applyStylePatch("fontStyle", styleOverride, {
          styleValueOverride: styleOverride,
          fontTokenFamilyOverride: requestedFamily,
          fontTokenOverride: tokenCandidate || suppliedFontToken
        });
        var readbackMatch = readbackMatchesExpectedFamily(requestedFamily);
        if (readbackMatch === false) {
          continue;
        }
        return true;
      }
    }
    return false;
  }

  if (normalizedStyleKey === "fontStyle" && extractedStyleValues && extractedStyleValues.fontFamily) {
    var expectedFamily = extractedStyleValues.fontFamily;
    var expectedFamilyKey = subcreator_visual_normalize_font_compare_key(expectedFamily);
    var suppliedStyleFontToken = extraOptions && typeof extraOptions.fontToken === "string" ? extraOptions.fontToken : "";
    var styleCandidates = [];

    function pushUniqueStyleCandidate(candidateValue) {
      // // Keep candidate style retries deterministic and unique.
      var text = subcreator_trim_string(String(candidateValue || ""));
      if (!text) {
        return;
      }
      var normalized = text.toLowerCase();
      for (var styleCandidateIndex = 0; styleCandidateIndex < styleCandidates.length; styleCandidateIndex += 1) {
        if (String(styleCandidates[styleCandidateIndex] || "").toLowerCase() === normalized) {
          return;
        }
      }
      styleCandidates.push(text);
    }

    pushUniqueStyleCandidate(normalizedStyleValue);
    if (extractedStyleValues.fontStyleOptions && typeof extractedStyleValues.fontStyleOptions.length === "number") {
      for (var styleOptionIndex = 0; styleOptionIndex < extractedStyleValues.fontStyleOptions.length; styleOptionIndex += 1) {
        pushUniqueStyleCandidate(extractedStyleValues.fontStyleOptions[styleOptionIndex]);
      }
    }
    pushUniqueStyleCandidate("Regular");
    pushUniqueStyleCandidate("Book");
    pushUniqueStyleCandidate("Medium");
    pushUniqueStyleCandidate("Roman");

    for (var styleRetryIndex = 0; styleRetryIndex < styleCandidates.length; styleRetryIndex += 1) {
      var candidateStyle = styleCandidates[styleRetryIndex];
      var styleTokenCandidates = subcreator_visual_build_font_token_candidates(
        expectedFamily,
        candidateStyle,
        suppliedStyleFontToken,
        ""
      );
      if (styleTokenCandidates.length < 1) {
        styleTokenCandidates.push("");
      }
      for (var styleTokenCandidateIndex = 0; styleTokenCandidateIndex < styleTokenCandidates.length; styleTokenCandidateIndex += 1) {
        var styleTokenCandidate = styleTokenCandidates[styleTokenCandidateIndex];
        var styleApplyConfig = {
          styleValueOverride: candidateStyle,
          fontTokenFamilyOverride: expectedFamily
        };
        if (styleTokenCandidate) {
          styleApplyConfig.fontTokenOverride = styleTokenCandidate;
        }
        if (!applyOnce(styleApplyConfig)) {
          continue;
        }
        if (typeof property.getValue !== "function") {
          return true;
        }
        try {
          var styleReadbackRaw = property.getValue();
          var styleReadback = subcreator_visual_extract_text_style_from_value(styleReadbackRaw);
          if (!styleReadback || !styleReadback.fontFamily) {
            return true;
          }
          var readbackFamilyKey = subcreator_visual_normalize_font_compare_key(styleReadback.fontFamily);
          if (!expectedFamilyKey || !readbackFamilyKey || readbackFamilyKey === expectedFamilyKey) {
            return true;
          }
        } catch (styleReadbackError) {
          return true;
        }
      }
    }
    return false;
  }

  return applyOnce({
    fontTokenOverride: extraOptions && typeof extraOptions.fontToken === "string" ? extraOptions.fontToken : ""
  });
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

function subcreator_visual_numeric_values_match(targetValue, readbackValue) {
  // // Compare numeric writes with a tiny tolerance so host readback can confirm that `0` really stuck.
  var target = Number(targetValue);
  var readback = Number(readbackValue);
  if (isNaN(target) || isNaN(readback)) {
    return false;
  }
  return Math.abs(target - readback) <= 0.0001;
}

function subcreator_visual_try_set_numeric_property(property, numericValue, debugLines, debugLabel) {
  // // Validate numeric writes with readback because some Premiere sliders silently coerce values like `0`.
  if (!property || typeof property.setValue !== "function") {
    return false;
  }

  var attempts = [
    { value: numericValue, useRefresh: true, label: "number_refresh" },
    { value: numericValue, useRefresh: false, label: "number_no_refresh" },
    { value: String(numericValue), useRefresh: true, label: "string_refresh" },
    { value: String(numericValue), useRefresh: false, label: "string_no_refresh" }
  ];

  for (var attemptIndex = 0; attemptIndex < attempts.length; attemptIndex += 1) {
    var attempt = attempts[attemptIndex];
    try {
      property.setValue(attempt.value, attempt.useRefresh);
    } catch (numericSetError) {
      continue;
    }

    if (typeof property.getValue !== "function") {
      return true;
    }

    try {
      var readbackValue = property.getValue();
      if (subcreator_visual_numeric_values_match(numericValue, readbackValue)) {
        if (attemptIndex > 0 && debugLines && debugLines.push) {
          debugLines.push(
            "numeric write fallback label=" +
              String(debugLabel || "") +
              " mode=" +
              attempt.label +
              " target=" +
              String(numericValue)
          );
        }
        return true;
      }

      if (attemptIndex === attempts.length - 1 && debugLines && debugLines.push) {
        debugLines.push(
          "numeric write mismatch label=" +
            String(debugLabel || "") +
            " target=" +
            String(numericValue) +
            " readback=" +
            String(readbackValue)
        );
      }
    } catch (numericReadbackError) {
      if (attemptIndex === 0) {
        return true;
      }
    }
  }

  return false;
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
    var mogrtItems = subcreator_collect_selected_mogrt_items(sequence);

    if (!mogrtItems.length) {
      return subcreator_ok({
        selectedCount: 0,
        editableCount: 0,
        properties: []
      });
    }

    var firstTrackItem = mogrtItems[0];
    var rawComponents = subcreator_get_mogrt_components_from_track_item(firstTrackItem);
    var components = [];
    var properties = [];
    var componentDebugEntries = [];
    var duplicateComponentDebug = [];
    var sequenceSize = subcreator_visual_read_sequence_dimensions();
    var seenComponentSignatures = {};
    subcreator_visual_reset_group_sequence_axis_preferences();
    subcreator_visual_reset_text_style_option_cache();
    for (var componentIndex = 0; componentIndex < rawComponents.length; componentIndex += 1) {
      var rawComponent = rawComponents[componentIndex];
      var componentGroupPath = rawComponents.length > 1 ? subcreator_visual_get_component_group_label(rawComponent, componentIndex) : "";
      var componentProperties = [];
      subcreator_collect_mogrt_visual_properties_recursive(
        rawComponent ? rawComponent.properties : null,
        subcreator_visual_build_component_path_prefix(componentIndex),
        componentGroupPath,
        componentProperties
      );

      componentProperties = subcreator_visual_filter_property_descriptors(componentProperties);
      var componentSignature = subcreator_visual_build_component_descriptor_signature(componentProperties);
      if (componentSignature && typeof seenComponentSignatures[componentSignature] !== "undefined") {
        // // Collapse duplicated AE components that expose the same visible controls twice in the host API.
        duplicateComponentDebug.push({
          index: componentIndex,
          duplicateOf: seenComponentSignatures[componentSignature],
          name: subcreator_trim_string(String((rawComponent && (rawComponent.displayName || rawComponent.name)) || "")) || "Component " + String(componentIndex),
          propertyCount: componentProperties.length
        });
        continue;
      }

      seenComponentSignatures[componentSignature] = componentIndex;
      components.push(rawComponent);
      componentDebugEntries.push({
        index: componentIndex,
        name: subcreator_trim_string(String((rawComponent && (rawComponent.displayName || rawComponent.name)) || "")) || "Component " + String(componentIndex),
        propertyCount: componentProperties.length
      });

      for (var componentPropertyIndex = 0; componentPropertyIndex < componentProperties.length; componentPropertyIndex += 1) {
        properties.push(componentProperties[componentPropertyIndex]);
      }
    }

    var debug = {
      sequenceWidth: sequenceSize.width,
      sequenceHeight: sequenceSize.height,
      rawComponentCount: rawComponents.length,
      componentCount: components.length,
      components: componentDebugEntries,
      duplicateComponents: duplicateComponentDebug,
      vectorCount: 0,
      colorCount: 0,
      selectCount: 0,
      sample: [],
      textStyleCandidates: []
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
        var samplePathState = subcreator_visual_parse_component_prefixed_path(item.path);
        var sampleEntry = {
          path: item.path,
          componentIndex: samplePathState.componentIndex,
          name: item.displayName,
          group: item.groupPath,
          kind: item.controlKind,
          value: item.value,
          vectorScale: item.vectorScale || null,
          vectorMode: item.vectorMode || null
        };

        if (item.controlKind === "color") {
          // // Include raw color API/value snapshots to troubleshoot host channel-order inconsistencies.
          var sampleResolved = subcreator_visual_resolve_property_from_track_item(firstTrackItem, item.path);
          if (sampleResolved && sampleResolved.property) {
            try {
              var sampleColorApiValue =
                typeof sampleResolved.property.getColorValue === "function"
                  ? sampleResolved.property.getColorValue()
                  : "<no getColorValue>";
              sampleEntry.colorApiRaw =
                typeof sampleColorApiValue === "string" ? sampleColorApiValue : JSON.stringify(sampleColorApiValue);
            } catch (sampleColorApiError) {
              sampleEntry.colorApiRaw = "<error " + String(sampleColorApiError) + ">";
            }

            try {
              var sampleRawValue =
                typeof sampleResolved.property.getValue === "function"
                  ? sampleResolved.property.getValue()
                  : "<no getValue>";
              sampleEntry.valueRaw = typeof sampleRawValue === "string" ? sampleRawValue : JSON.stringify(sampleRawValue);
            } catch (sampleRawError) {
              sampleEntry.valueRaw = "<error " + String(sampleRawError) + ">";
            }
          }
        }

        debug.sample.push(sampleEntry);
      }
    }

    for (var debugComponentIndex = 0; debugComponentIndex < componentDebugEntries.length; debugComponentIndex += 1) {
      var debugComponentEntry = componentDebugEntries[debugComponentIndex];
      var debugComponent = rawComponents[debugComponentEntry.index];
      if (!debugComponentEntry || !debugComponent || !debugComponent.properties) {
        continue;
      }

      subcreator_collect_text_style_debug_candidates(
        debugComponent.properties,
        subcreator_visual_build_component_path_prefix(debugComponentEntry.index),
        rawComponents.length > 1 ? subcreator_visual_get_component_group_label(debugComponent, debugComponentEntry.index) : "",
        debug.textStyleCandidates,
        20
      );
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

function subcreator_get_selected_mogrt_count() {
  // // Return current selected MOGRT clip count for panel-side progress rendering.
  try {
    if (!app || !app.project || !app.project.activeSequence) {
      return subcreator_error("No active sequence in Premiere.");
    }

    var sequence = app.project.activeSequence;
    var mogrtItems = subcreator_collect_selected_mogrt_items(sequence);
    return subcreator_ok({
      selectedCount: mogrtItems.length
    });
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_list_selected_mogrt_text_items() {
  // // Return selected MOGRT clips as editable subtitle text blocks for the Text tab.
  try {
    if (!app || !app.project || !app.project.activeSequence) {
      return subcreator_error("No active sequence in Premiere.");
    }

    var sequence = app.project.activeSequence;
    var sequenceIdentity = subcreator_get_sequence_identity(sequence);
    var mogrtItems = subcreator_collect_resolved_selected_mogrt_items(sequence);
    if (!mogrtItems.length) {
      return subcreator_ok({
        selectedCount: 0,
        projectDocumentId: sequenceIdentity.projectDocumentId,
        projectPath: sequenceIdentity.projectPath,
        sequenceID: sequenceIdentity.sequenceID,
        sequenceName: sequenceIdentity.sequenceName,
        signature: "",
        items: []
      });
    }

    var firstTrackIndex = subcreator_find_track_item_video_track_index(sequence, mogrtItems[0]);
    var sameTrack = true;
    var items = [];

    for (var itemIndex = 0; itemIndex < mogrtItems.length; itemIndex += 1) {
      var trackItem = mogrtItems[itemIndex];
      var trackIndex = subcreator_find_track_item_video_track_index(sequence, trackItem);
      if (trackIndex !== firstTrackIndex) {
        sameTrack = false;
      }

      items.push({
        selectionIndex: itemIndex,
        videoTrackIndex: trackIndex,
        startSeconds: subcreator_to_seconds(trackItem.start || trackItem.inPoint || trackItem.startTime),
        endSeconds: subcreator_to_seconds(trackItem.end || trackItem.outPoint || trackItem.endTime),
        text: subcreator_trim_string(String(subcreator_extract_text_from_mogrt_item(trackItem) || "").replace(/\s+/g, " ")),
        clipName: subcreator_trim_string(String((trackItem.projectItem && trackItem.projectItem.name) || trackItem.name || "MOGRT"))
      });
    }

    return subcreator_ok({
      selectedCount: mogrtItems.length,
      sameTrack: sameTrack,
      videoTrackIndex: firstTrackIndex,
      projectDocumentId: sequenceIdentity.projectDocumentId,
      projectPath: sequenceIdentity.projectPath,
      sequenceID: sequenceIdentity.sequenceID,
      sequenceName: sequenceIdentity.sequenceName,
      signature: subcreator_build_selected_mogrt_text_signature(sequence, mogrtItems),
      items: items
    });
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_apply_selected_mogrt_text_items(payloadEncoded) {
  // // Rebuild selected subtitle MOGRT clips from edited text blocks, staying on the source track only when the insertion is safe.
  try {
    if (!app || !app.project || !app.project.activeSequence) {
      return subcreator_error("No active sequence in Premiere.");
    }

    var sequence = app.project.activeSequence;
    var textApplyFrameDurationSeconds = subcreator_get_sequence_frame_duration_seconds(sequence);
    var sequenceIdentity = subcreator_get_sequence_identity(sequence);
    var decodedPayload = subcreator_decode_payload(payloadEncoded || "");
    var payload = JSON.parse(decodedPayload || "{}");
    var editedItems = payload && payload.items && typeof payload.items.length === "number" ? payload.items : [];
    var currentSelection = subcreator_collect_resolved_selected_mogrt_items(sequence);
    var currentSignature = subcreator_build_selected_mogrt_text_signature(sequence, currentSelection);
    var expectedSignature = subcreator_trim_string(String(payload.selectionSignature || ""));

    if (!currentSelection.length) {
      return subcreator_ok({
        selectedCount: 0,
        rebuiltCount: 0,
        failedCount: 0,
        debug: ["text_apply selection empty"]
      });
    }

    if (expectedSignature && currentSignature !== expectedSignature) {
      return subcreator_error("Selection changed since last read. Reload the Text tab selection before applying.");
    }

    var targetTrackIndex = subcreator_find_track_item_video_track_index(sequence, currentSelection[0]);
    if (targetTrackIndex < 0) {
      return subcreator_error("Unable to resolve selected MOGRT track.");
    }
    var sourceTrackIndex = targetTrackIndex;

    for (var selectedIndex = 0; selectedIndex < currentSelection.length; selectedIndex += 1) {
      if (subcreator_find_track_item_video_track_index(sequence, currentSelection[selectedIndex]) !== targetTrackIndex) {
        return subcreator_error("Text tab currently supports selected MOGRT clips on one video track only.");
      }
    }

    var track = sequence.videoTracks ? sequence.videoTracks[targetTrackIndex] : null;
    if (!track) {
      return subcreator_error("Unable to access target video track.");
    }

    var replaceSelectionStartIndex = Number(payload.replaceSelectionStartIndex);
    var replaceSelectionEndIndex = Number(payload.replaceSelectionEndIndex);
    if (isNaN(replaceSelectionStartIndex) || replaceSelectionStartIndex < 0) {
      replaceSelectionStartIndex = 0;
    }
    if (isNaN(replaceSelectionEndIndex) || replaceSelectionEndIndex >= currentSelection.length) {
      replaceSelectionEndIndex = currentSelection.length - 1;
    }
    if (replaceSelectionEndIndex < replaceSelectionStartIndex) {
      replaceSelectionEndIndex = replaceSelectionStartIndex;
    }

    var selectionItemsToReplace = [];
    var untouchedSelectedTrackItems = [];
    for (var currentIndex = 0; currentIndex < currentSelection.length; currentIndex += 1) {
      if (currentIndex >= replaceSelectionStartIndex && currentIndex <= replaceSelectionEndIndex) {
        selectionItemsToReplace.push(currentSelection[currentIndex]);
      } else {
        untouchedSelectedTrackItems.push(currentSelection[currentIndex]);
      }
    }

    if (selectionItemsToReplace.length < 1) {
      return subcreator_error("No selected subtitle range available for text rebuild.");
    }

    var sourceSnapshots = [];
    for (var snapshotIndex = 0; snapshotIndex < currentSelection.length; snapshotIndex += 1) {
      sourceSnapshots.push({
        visualChanges: subcreator_build_visual_clone_changes_from_track_item(currentSelection[snapshotIndex]),
        textChanges: subcreator_build_text_clone_changes_from_track_item(currentSelection[snapshotIndex])
      });
    }

    var fallbackOptions = payload && payload.options && typeof payload.options === "object" ? payload.options : {};
    var fallbackMogrtPath = subcreator_resolve_mogrt_path(fallbackOptions);
    var fallbackPathCandidates = subcreator_build_mogrt_path_candidates(fallbackMogrtPath);
    var debugLines = [];
    var rebuiltCount = 0;
    var failedCount = 0;
    var clonedStyleUpdates = 0;
    var clonedStyleFailures = 0;
    var clonedTextUpdates = 0;
    var clonedTextFailures = 0;
    var durationAdjusted = 0;
    var selectedTrackItems = [];

    debugLines.push("text_apply selected=" + String(currentSelection.length) + " edited=" + String(editedItems.length));
    debugLines.push(
      "text_apply replace_range=" + String(replaceSelectionStartIndex) + "-" + String(replaceSelectionEndIndex)
    );
    debugLines.push("text_apply source_track=" + String(targetTrackIndex));

    if (fallbackPathCandidates.length < 1) {
      return subcreator_error(
        "Sub Creator could not resolve the source MOGRT file for a safe text rebuild. Select the matching template in the gallery, then retry."
      );
    }

    var replaceRangeEndSeconds = Number.NEGATIVE_INFINITY;
    for (var replaceIndex = 0; replaceIndex < selectionItemsToReplace.length; replaceIndex += 1) {
      replaceRangeEndSeconds = Math.max(
        replaceRangeEndSeconds,
        subcreator_to_seconds(
          selectionItemsToReplace[replaceIndex] &&
            (selectionItemsToReplace[replaceIndex].end ||
              selectionItemsToReplace[replaceIndex].outPoint ||
              selectionItemsToReplace[replaceIndex].endTime)
        )
      );
    }

    var followingClipConflict = subcreator_find_text_rebuild_following_clip(track, selectionItemsToReplace, replaceRangeEndSeconds);
    if (followingClipConflict) {
      var laterClipFallbackTrackInfo = subcreator_get_or_create_video_track_above_index(sequence, targetTrackIndex);
      if (laterClipFallbackTrackInfo && laterClipFallbackTrackInfo.index >= 0) {
        targetTrackIndex = laterClipFallbackTrackInfo.index;
        track = sequence.videoTracks[targetTrackIndex];
        debugLines.push(
          "text_apply fallback_track=" +
            String(targetTrackIndex) +
            " reason=later_clip clip=" +
            String(followingClipConflict.clipName || "") +
            " created=" +
            (laterClipFallbackTrackInfo.created ? "true" : "false")
        );
      } else {
        return subcreator_error(
          "Text editor apply detected later clips on the source track and Sub Creator could not create a safe fallback video track above."
        );
      }
    }

    // Keep the full current selection exempt from overlap checks so untouched selected subtitles do not trigger a false fallback track.
    var overlapConflict = subcreator_find_text_rebuild_overlap(track, currentSelection, editedItems, selectionItemsToReplace);
    if (overlapConflict) {
      var fallbackTrackInfo = subcreator_get_or_create_video_track_above_index(sequence, targetTrackIndex);
      if (fallbackTrackInfo && fallbackTrackInfo.index >= 0) {
        targetTrackIndex = fallbackTrackInfo.index;
        track = sequence.videoTracks[targetTrackIndex];
        debugLines.push(
          "text_apply fallback_track=" +
            String(targetTrackIndex) +
            " reason=overlap clip=" +
            String(overlapConflict.clipName || "") +
            " created=" +
            (fallbackTrackInfo.created ? "true" : "false")
        );
      } else {
        debugLines.push(
          "text_apply overlap_abort clip=" +
            String(overlapConflict.clipName || "") +
            " range=" +
            String(Math.round(Number(overlapConflict.startSeconds || 0) * 1000)) +
            "-" +
            String(Math.round(Number(overlapConflict.endSeconds || 0) * 1000))
        );
        return subcreator_error(
          "Text editor apply would overlap non-selected clip '" +
            String(overlapConflict.clipName || "Clip") +
            "' on the same track, and Sub Creator could not create a safe fallback video track above."
        );
      }
    }
    debugLines.push("text_apply rebuild_track=" + String(targetTrackIndex));

    for (var removeIndex = selectionItemsToReplace.length - 1; removeIndex >= 0; removeIndex -= 1) {
      if (!subcreator_remove_track_item_without_ripple(selectionItemsToReplace[removeIndex])) {
        failedCount += 1;
        debugLines.push("text_apply remove_failed index=" + String(removeIndex));
      }
    }

    for (var editedIndex = 0; editedIndex < editedItems.length; editedIndex += 1) {
      var editedItem = editedItems[editedIndex] || {};
      var textValue = subcreator_trim_string(
        String(editedItem.text || "")
          .replace(/\r\n?/g, "\n")
          .replace(/[ \t]+/g, " ")
          .replace(/\n{3,}/g, "\n\n")
      );
      var startSeconds = Number(editedItem.startSeconds);
      var endSeconds = Number(editedItem.endSeconds);
      var nextEditedItem = editedIndex + 1 < editedItems.length ? editedItems[editedIndex + 1] || {} : null;
      var originalStartSeconds = startSeconds;
      var originalEndSeconds = endSeconds;
      startSeconds = subcreator_snap_seconds_to_nearest_frame(startSeconds, textApplyFrameDurationSeconds);
      endSeconds = subcreator_snap_seconds_to_nearest_frame(endSeconds, textApplyFrameDurationSeconds);
      var nextEditedStartSeconds = nextEditedItem
        ? subcreator_snap_seconds_to_nearest_frame(Number(nextEditedItem.startSeconds), textApplyFrameDurationSeconds)
        : NaN;
      endSeconds = subcreator_snap_nearby_caption_end(
        endSeconds,
        nextEditedStartSeconds,
        textApplyFrameDurationSeconds
      );
      if (endSeconds <= startSeconds) {
        endSeconds = startSeconds + textApplyFrameDurationSeconds;
      }
      if (startSeconds !== originalStartSeconds || endSeconds !== originalEndSeconds) {
        debugLines.push(
          "text_apply frame_aligned=" +
            String(originalStartSeconds) +
            "-" +
            String(originalEndSeconds) +
            " -> " +
            String(startSeconds) +
            "-" +
            String(endSeconds)
        );
      }
      var sourceSelectionIndex = Number(editedItem.sourceSelectionIndex);
      var itemMogrtPathOverride = subcreator_trim_string(String(editedItem.mogrtPathOverride || ""));
      var itemSkipTextApply = Boolean(editedItem.skipTextApply);
      if (isNaN(sourceSelectionIndex) || sourceSelectionIndex < 0 || sourceSelectionIndex >= sourceSnapshots.length) {
        sourceSelectionIndex = 0;
      }

      if (!textValue || isNaN(startSeconds) || isNaN(endSeconds) || endSeconds <= startSeconds) {
        failedCount += 1;
        debugLines.push("text_apply invalid_item index=" + String(editedIndex));
        continue;
      }

      var sourceSnapshot = sourceSnapshots[sourceSelectionIndex] || sourceSnapshots[0] || {
        visualChanges: []
      };
      var insertedTrackItem = null;
      var itemPathCandidates = itemMogrtPathOverride ? subcreator_build_mogrt_path_candidates(itemMogrtPathOverride) : fallbackPathCandidates;

      if (itemPathCandidates.length > 0) {
        var importAttempt = subcreator_try_import_mogrt(sequence, itemPathCandidates, startSeconds, targetTrackIndex, 0);
        insertedTrackItem = importAttempt.trackItem;
        if (insertedTrackItem) {
          debugLines.push(
            "text_apply import_fallback path=" + String(importAttempt.usedPath || "") + " mode=" + String(importAttempt.usedTimeMode || "")
          );
        }
      }

      if (!insertedTrackItem) {
        failedCount += 1;
        debugLines.push("text_apply insert_failed index=" + String(editedIndex));
        continue;
      }

      if (sourceSnapshot.visualChanges && sourceSnapshot.visualChanges.length > 0) {
        var cloneStats = subcreator_apply_visual_changes_to_track_item(insertedTrackItem, sourceSnapshot.visualChanges, debugLines);
        clonedStyleUpdates += cloneStats.updatedCount;
        clonedStyleFailures += cloneStats.failedCount;
      }

      var rebuiltDurationStyleConfig = subcreator_clone_style_config_with_clip_duration(null, endSeconds - startSeconds);
      var rebuiltDurationStats = subcreator_try_set_mogrt_controls(
        insertedTrackItem,
        textValue,
        "",
        rebuiltDurationStyleConfig,
        [],
        true,
        debugLines,
        "text_apply duration_control index=" + String(editedIndex)
      );
      if (rebuiltDurationStats && rebuiltDurationStats.layoutUpdates > 0) {
        debugLines.push("text_apply duration_control_updates=" + String(rebuiltDurationStats.layoutUpdates));
      }

      var textUpdateCount = 0;
      if (!itemSkipTextApply && sourceSnapshot.textChanges && sourceSnapshot.textChanges.length > 0) {
        var textCloneStats = subcreator_apply_text_clone_changes_to_track_item(
          insertedTrackItem,
          sourceSnapshot.textChanges,
          textValue,
          debugLines
        );
        clonedTextUpdates += textCloneStats.updatedCount;
        clonedTextFailures += textCloneStats.failedCount;
        textUpdateCount = textCloneStats.updatedCount;
      }

      if (!itemSkipTextApply && textUpdateCount < 1) {
        var textStats = subcreator_try_set_mogrt_controls(insertedTrackItem, textValue, "", null, [], false, null, null);
        textUpdateCount = textStats && textStats.textUpdates ? textStats.textUpdates : 0;
      }

      if (itemSkipTextApply) {
        debugLines.push("text_apply text_baked index=" + String(editedIndex));
      } else if (textUpdateCount < 1) {
        debugLines.push("text_apply text_update_missing index=" + String(editedIndex));
      }

      var textApplyDurationResult = subcreator_try_razor_mogrt_duration(
        sequence,
        insertedTrackItem,
        targetTrackIndex,
        startSeconds,
        endSeconds,
        debugLines,
        "text_apply index=" + String(editedIndex)
      );
      if (textApplyDurationResult.applied) {
        durationAdjusted += 1;
        insertedTrackItem = textApplyDurationResult.trackItem;
      } else {
        debugLines.push("text_apply duration_failed index=" + String(editedIndex));
      }
      selectedTrackItems.push(insertedTrackItem);
      rebuiltCount += 1;
    }

    var selectionAfterApply = subcreator_sort_track_items_by_time(untouchedSelectedTrackItems.concat(selectedTrackItems));
    for (var selectIndex = 0; selectIndex < selectionAfterApply.length; selectIndex += 1) {
      subcreator_try_select_track_item(selectionAfterApply[selectIndex], selectIndex === 0);
    }

    var refreshTriggered = subcreator_force_sequence_visual_refresh(sequence);
    debugLines.push("text_apply cloned_style_updates=" + String(clonedStyleUpdates));
    debugLines.push("text_apply cloned_style_failures=" + String(clonedStyleFailures));
    debugLines.push("text_apply cloned_text_updates=" + String(clonedTextUpdates));
    debugLines.push("text_apply cloned_text_failures=" + String(clonedTextFailures));
    debugLines.push("text_apply duration_adjusted=" + String(durationAdjusted));
    debugLines.push("text_apply ui_refresh=" + (refreshTriggered ? "forced" : "not_available"));

    return subcreator_ok({
      selectedCount: currentSelection.length,
      rebuiltCount: rebuiltCount,
      failedCount: failedCount,
      selectionSignature: subcreator_build_selected_mogrt_text_signature(sequence, selectionAfterApply),
      sourceTrackIndex: sourceTrackIndex,
      rebuildTrackIndex: targetTrackIndex,
      projectDocumentId: sequenceIdentity.projectDocumentId,
      projectPath: sequenceIdentity.projectPath,
      sequenceID: sequenceIdentity.sequenceID,
      sequenceName: sequenceIdentity.sequenceName,
      debug: debugLines
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
    var mogrtItems = subcreator_collect_selected_mogrt_items(sequence);

    if (!mogrtItems.length) {
      return subcreator_ok({
        selectedCount: 0,
        updatedCount: 0,
        failedCount: 0,
        processedClipCount: 0
      });
    }

    var clipStartIndex = Number(payload.clipStartIndex);
    if (isNaN(clipStartIndex) || clipStartIndex < 0) {
      clipStartIndex = 0;
    } else {
      clipStartIndex = Math.floor(clipStartIndex);
    }
    if (clipStartIndex > mogrtItems.length) {
      clipStartIndex = mogrtItems.length;
    }

    var clipEndIndex = Number(payload.clipEndIndex);
    if (isNaN(clipEndIndex)) {
      clipEndIndex = mogrtItems.length;
    } else {
      clipEndIndex = Math.floor(clipEndIndex);
    }
    if (clipEndIndex < clipStartIndex) {
      clipEndIndex = clipStartIndex;
    }
    if (clipEndIndex > mogrtItems.length) {
      clipEndIndex = mogrtItems.length;
    }

    var processedClipCount = Math.max(0, clipEndIndex - clipStartIndex);

    var updatedCount = 0;
    var failedCount = 0;
    var debugLines = [];
    var applySequenceSize = subcreator_visual_read_sequence_dimensions();
    debugLines.push("sequence=" + applySequenceSize.width + "x" + applySequenceSize.height);
    debugLines.push(
      "clip_range=" + String(clipStartIndex) + "-" + String(clipEndIndex) + " selected=" + String(mogrtItems.length)
    );

    for (var clipIndex = clipStartIndex; clipIndex < clipEndIndex; clipIndex += 1) {
      var clip = mogrtItems[clipIndex];
      var clipComponents = subcreator_get_mogrt_components_from_track_item(clip);
      if (clipComponents.length < 1) {
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
        var fontToken = subcreator_trim_string(String(change.fontToken || ""));
        var vectorScale = null;
        if (change.vectorScale && Object.prototype.toString.call(change.vectorScale) === "[object Array]") {
          vectorScale = change.vectorScale;
        }
        if (!path) {
          failedCount += 1;
          continue;
        }

        var resolvedProperty = subcreator_visual_resolve_property_from_track_item(clip, resolvedPath);
        var property = resolvedProperty ? resolvedProperty.property : null;
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
          (valueType === "number" && Number(value) === 0) ||
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
              (fontToken ? " fontToken=" + fontToken : "") +
              (vectorScale ? " scale=" + String(vectorScale) : "")
          );
        }

        if (virtualTextStyleTarget) {
          try {
            applied = subcreator_try_set_mogrt_text_style_property(property, virtualTextStyleTarget.styleKey, value, {
              fontToken: fontToken
            });
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
        } else if ((controlKind === "slider" || controlKind === "number") && valueType === "number") {
          try {
            applied = subcreator_visual_try_set_numeric_property(
              property,
              Number(value),
              debugLines,
              path + " name=" + displayName
            );
          } catch (numericError) {}
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
          if (virtualTextStyleTarget) {
            try {
              var textStyleReadbackRaw = typeof property.getValue === "function" ? property.getValue() : "";
              var textStyleReadback = subcreator_visual_extract_text_style_from_value(textStyleReadbackRaw);
              if (textStyleReadback) {
                debugLines.push(
                  "textstyle readback family=" +
                    String(textStyleReadback.fontFamily || "<none>") +
                    " style=" +
                    String(textStyleReadback.fontStyle || "<none>") +
                    " token=" +
                    String(textStyleReadback.fontToken || "<none>") +
                    " size=" +
                    String(textStyleReadback.fontSize || "<none>")
                );
              } else {
                debugLines.push("textstyle readback unavailable");
              }
            } catch (textReadbackError) {
              debugLines.push("textstyle readback failed: " + String(textReadbackError));
            }
          }
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
      processedClipCount: processedClipCount,
      clipStartIndex: clipStartIndex,
      clipEndIndex: clipEndIndex,
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

function subcreator_binary_string_to_byte_array(binaryValue) {
  // // Convert ExtendScript binary-like strings into byte arrays so Premiere flatbuffer text payloads can be inspected safely.
  var rawText = String(binaryValue || "");
  var bytes = [];
  for (var index = 0; index < rawText.length; index += 1) {
    bytes.push(rawText.charCodeAt(index) & 255);
  }
  return bytes;
}

function subcreator_byte_array_to_binary_string(bytes) {
  // // Rebuild a binary-safe ExtendScript string from byte values without truncating embedded null bytes.
  if (!bytes || typeof bytes.length !== "number" || bytes.length < 1) {
    return "";
  }

  var result = "";
  var chunkSize = 8192;
  for (var index = 0; index < bytes.length; index += chunkSize) {
    var slice = bytes.slice(index, index + chunkSize);
    result += String.fromCharCode.apply(null, slice);
  }
  return result;
}

function subcreator_utf8_text_to_byte_array(textValue) {
  // // Encode user text as UTF-8 bytes so Premiere-authored Source Text payloads keep their document styling blob intact.
  var encoded = "";
  try {
    encoded = unescape(encodeURIComponent(String(textValue || "")));
  } catch (encodeError) {
    encoded = String(textValue || "");
  }
  return subcreator_binary_string_to_byte_array(encoded);
}

function subcreator_utf8_byte_array_to_text(bytes) {
  // // Decode UTF-8 byte slices extracted from Premiere text-document buffers back into readable subtitle text.
  if (!bytes || typeof bytes.length !== "number" || bytes.length < 1) {
    return "";
  }

  try {
    return decodeURIComponent(escape(subcreator_byte_array_to_binary_string(bytes)));
  } catch (decodeError) {
    var fallback = "";
    for (var index = 0; index < bytes.length; index += 1) {
      fallback += String.fromCharCode(bytes[index] & 255);
    }
    return fallback;
  }
}

function subcreator_is_probably_binary_text_payload(rawValue) {
  // // Detect Premiere text-document blobs so the host does not accidentally replace them with plain strings and lose styling.
  if (typeof rawValue !== "string" || rawValue.length < 8 || rawValue.indexOf("{") !== -1) {
    return false;
  }

  var controlByteCount = 0;
  for (var index = 0; index < rawValue.length; index += 1) {
    var code = rawValue.charCodeAt(index);
    if (code === 0) {
      return true;
    }
    if (code < 32 && code !== 9 && code !== 10 && code !== 13) {
      controlByteCount += 1;
      if (controlByteCount >= 4) {
        return true;
      }
    }
  }

  return false;
}

function subcreator_score_binary_text_candidate(textValue, offset, totalLength, hasOnlyZeroPaddingAfter) {
  // // Prefer later and more sentence-like strings so flatbuffer scans pick the visible subtitle text instead of font names.
  var normalized = String(textValue || "");
  if (!normalized) {
    return -9999;
  }

  var score = 0;
  score += Math.max(Number(offset || 0), 0) / Math.max(Number(totalLength || 1), 1) * 5;
  score += normalized.split(/\s+/).filter(Boolean).length * 2;
  score += Math.min(normalized.length, 120) / 20;

  if (/[.,!?;:]/.test(normalized)) {
    score += 1;
  }

  if (/[\r\n]/.test(normalized)) {
    score += 1;
  }

  if (hasOnlyZeroPaddingAfter) {
    score += 3;
  }

  if (normalized.length < 3) {
    score -= 10;
  }

  return score;
}

function subcreator_find_binary_text_payload_candidate(rawValue) {
  // // Locate the likely visible text string inside a Premiere flatbuffer-style text payload so only the text bytes change.
  if (!subcreator_is_probably_binary_text_payload(rawValue)) {
    return null;
  }

  var bytes = subcreator_binary_string_to_byte_array(rawValue);
  if (!bytes || bytes.length < 8) {
    return null;
  }

  var bestCandidate = null;
  for (var index = 0; index <= bytes.length - 5; index += 1) {
    var byteLength =
      (bytes[index] & 255) |
      ((bytes[index + 1] & 255) << 8) |
      ((bytes[index + 2] & 255) << 16) |
      ((bytes[index + 3] & 255) << 24);

    if (byteLength < 1 || byteLength > 2000) {
      continue;
    }

    var textStart = index + 4;
    var textEnd = textStart + byteLength;
    if (textEnd >= bytes.length || bytes[textEnd] !== 0) {
      continue;
    }

    var candidateBytes = bytes.slice(textStart, textEnd);
    var candidateText = subcreator_utf8_byte_array_to_text(candidateBytes);
    if (!candidateText) {
      continue;
    }

    var printableCount = 0;
    for (var printableIndex = 0; printableIndex < candidateBytes.length; printableIndex += 1) {
      var candidateByte = candidateBytes[printableIndex] & 255;
      if ((candidateByte >= 32 && candidateByte <= 126) || candidateByte === 9 || candidateByte === 10 || candidateByte === 13 || candidateByte >= 128) {
        printableCount += 1;
      }
    }

    if (printableCount / candidateBytes.length < 0.8) {
      continue;
    }

    var zeroPaddingLength = 0;
    var suffixHasOnlyZeroPadding = true;
    for (var suffixIndex = textEnd + 1; suffixIndex < bytes.length; suffixIndex += 1) {
      if (bytes[suffixIndex] !== 0) {
        suffixHasOnlyZeroPadding = false;
        break;
      }
      zeroPaddingLength += 1;
    }

    var score = subcreator_score_binary_text_candidate(candidateText, index, bytes.length, suffixHasOnlyZeroPadding);
    if (!bestCandidate || score > bestCandidate.score) {
      bestCandidate = {
        score: score,
        offset: index,
        byteLength: byteLength,
        text: String(candidateText || "").replace(/\r/g, "\n"),
        hasOnlyZeroPaddingAfter: suffixHasOnlyZeroPadding,
        zeroPaddingLength: zeroPaddingLength
      };
    }
  }

  return bestCandidate;
}

function subcreator_try_append_flatbuffer_text_string(sourceBytes, stringOffset, replacementBytes) {
  // // Some Premiere text payloads cannot be resized in place; append the new string and retarget existing string pointers instead.
  if (!sourceBytes || typeof sourceBytes.length !== "number" || stringOffset < 4 || !replacementBytes) {
    return null;
  }

  var pointerOffsets = [];
  for (var offset = 0; offset <= stringOffset - 4; offset += 4) {
    var pointerValue =
      (sourceBytes[offset] & 255) |
      ((sourceBytes[offset + 1] & 255) << 8) |
      ((sourceBytes[offset + 2] & 255) << 16) |
      ((sourceBytes[offset + 3] & 255) << 24);
    if (pointerValue > 0 && offset + pointerValue === stringOffset) {
      pointerOffsets.push(offset);
    }
  }

  if (!pointerOffsets.length) {
    return null;
  }

  var patchedBytes = sourceBytes.slice(0);
  while (patchedBytes.length % 4 !== 0) {
    patchedBytes.push(0);
  }

  var appendedStringOffset = patchedBytes.length;
  patchedBytes.push(replacementBytes.length & 255);
  patchedBytes.push((replacementBytes.length >> 8) & 255);
  patchedBytes.push((replacementBytes.length >> 16) & 255);
  patchedBytes.push((replacementBytes.length >> 24) & 255);
  for (var byteIndex = 0; byteIndex < replacementBytes.length; byteIndex += 1) {
    patchedBytes.push(replacementBytes[byteIndex] & 255);
  }
  patchedBytes.push(0);
  while (patchedBytes.length % 4 !== 0) {
    patchedBytes.push(0);
  }

  for (var pointerIndex = 0; pointerIndex < pointerOffsets.length; pointerIndex += 1) {
    var pointerOffset = pointerOffsets[pointerIndex];
    var relativeOffset = appendedStringOffset - pointerOffset;
    if (relativeOffset < 1) {
      return null;
    }
    patchedBytes[pointerOffset] = relativeOffset & 255;
    patchedBytes[pointerOffset + 1] = (relativeOffset >> 8) & 255;
    patchedBytes[pointerOffset + 2] = (relativeOffset >> 16) & 255;
    patchedBytes[pointerOffset + 3] = (relativeOffset >> 24) & 255;
  }

  return patchedBytes;
}

function subcreator_try_patch_binary_text_payload(rawValue, textValue) {
  // // Rewrite only the terminal UTF-8 text segment of Premiere Source Text blobs so font/style data stays untouched.
  var candidate = subcreator_find_binary_text_payload_candidate(rawValue);
  if (!candidate) {
    return "";
  }

  var replacementBytes = subcreator_utf8_text_to_byte_array(subcreator_normalize_caption_text(textValue));
  var sourceBytes = subcreator_binary_string_to_byte_array(rawValue);
  if (!candidate.hasOnlyZeroPaddingAfter && replacementBytes.length !== candidate.byteLength) {
    var retargetedBytes = subcreator_try_append_flatbuffer_text_string(sourceBytes, candidate.offset, replacementBytes);
    if (retargetedBytes) {
      return subcreator_byte_array_to_binary_string(retargetedBytes);
    }
    return "";
  }

  var replacementLengthOffset = candidate.offset;
  var replacementTextOffset = candidate.offset + 4;
  var patchedBytes = sourceBytes.slice(0, replacementLengthOffset);

  patchedBytes.push(replacementBytes.length & 255);
  patchedBytes.push((replacementBytes.length >> 8) & 255);
  patchedBytes.push((replacementBytes.length >> 16) & 255);
  patchedBytes.push((replacementBytes.length >> 24) & 255);

  for (var byteIndex = 0; byteIndex < replacementBytes.length; byteIndex += 1) {
    patchedBytes.push(replacementBytes[byteIndex] & 255);
  }

  patchedBytes.push(0);

  if (candidate.hasOnlyZeroPaddingAfter) {
    for (var zeroIndex = 0; zeroIndex < candidate.zeroPaddingLength; zeroIndex += 1) {
      patchedBytes.push(0);
    }
  } else {
    for (var padIndex = 0; padIndex < candidate.byteLength - replacementBytes.length; padIndex += 1) {
      patchedBytes.push(0);
    }

    for (var suffixIndex = replacementTextOffset + candidate.byteLength + 1; suffixIndex < sourceBytes.length; suffixIndex += 1) {
      patchedBytes.push(sourceBytes[suffixIndex] & 255);
    }
  }

  return subcreator_byte_array_to_binary_string(patchedBytes);
}

function subcreator_decode_template_text_payloads(rawPayloads) {
  // // Decode CEP-provided Premiere template payloads once per apply so inserted clips can reuse the original Source Text document blob.
  if (!rawPayloads || typeof rawPayloads.length !== "number") {
    return [];
  }

  var decoded = [];
  for (var index = 0; index < rawPayloads.length; index += 1) {
    var payload = rawPayloads[index] || {};
    var sourcePayloadBase64 = subcreator_trim_string(String(payload.sourcePayloadBase64 || ""));
    if (!sourcePayloadBase64) {
      continue;
    }

    decoded.push({
      displayName: subcreator_trim_string(String(payload.displayName || "")),
      initialText: subcreator_trim_string(String(payload.initialText || "")).replace(/\r/g, "\n"),
      rawValue: subcreator_decode_base64_to_binary_string(sourcePayloadBase64)
    });
  }

  return decoded;
}

function subcreator_create_template_text_payload_state(templateTextPayloads) {
  // // Track per-display-name payload consumption so templates with multiple text layers can reuse payloads in traversal order.
  return {
    items: templateTextPayloads || [],
    usageByDisplayName: {}
  };
}

function subcreator_resolve_template_text_payload_raw_value(payloadState, displayName) {
  // // Prefer a matching template payload over the live property raw value when Premiere only exposes a destructive plain-string shorthand.
  if (!payloadState || !payloadState.items || !payloadState.items.length) {
    return "";
  }

  var normalizedDisplayName = subcreator_trim_string(String(displayName || "")).toLowerCase();
  var usageKey = normalizedDisplayName || "__fallback__";
  var usageCount = Number(payloadState.usageByDisplayName[usageKey] || 0);
  var matchedPayloads = [];

  for (var index = 0; index < payloadState.items.length; index += 1) {
    var candidate = payloadState.items[index];
    if (!candidate || !candidate.rawValue) {
      continue;
    }

    var candidateDisplayName = subcreator_trim_string(String(candidate.displayName || "")).toLowerCase();
    if (!normalizedDisplayName || candidateDisplayName === normalizedDisplayName) {
      matchedPayloads.push(candidate);
    }
  }

  if (!matchedPayloads.length && payloadState.items.length === 1) {
    matchedPayloads = payloadState.items.slice(0);
  }

  if (!matchedPayloads.length) {
    return "";
  }

  var resolvedPayload = matchedPayloads[Math.min(usageCount, matchedPayloads.length - 1)];
  payloadState.usageByDisplayName[usageKey] = usageCount + 1;
  return String(resolvedPayload.rawValue || "");
}

function subcreator_extract_text_from_property_value(rawValue) {
  // // Convert property values (plain/object/JSON string) into readable caption text.
  if (rawValue === undefined || rawValue === null) {
    return "";
  }

  if (typeof rawValue === "string") {
    var rawText = String(rawValue || "");
    var binaryCandidate = subcreator_find_binary_text_payload_candidate(rawText);
    if (binaryCandidate && binaryCandidate.text) {
      return String(binaryCandidate.text || "");
    }

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

  var components = subcreator_get_mogrt_components_from_track_item(item);
  for (var componentIndex = 0; componentIndex < components.length; componentIndex += 1) {
    var component = components[componentIndex];
    if (!component || !component.properties) {
      continue;
    }

    var extracted = subcreator_extract_text_from_component_properties(component.properties);
    if (extracted) {
      return extracted;
    }
  }

  return "";
}

function subcreator_is_non_textual_mogrt_label(text) {
  // // Reject generic component labels that Premiere exposes but that are not actual subtitle content.
  var normalized = subcreator_trim_string(String(text || "")).toLowerCase();
  return (
    normalized === "graphic parameters" ||
    normalized === "graphics parameters" ||
    normalized === "parametres graphiques" ||
    normalized === "essential graphics"
  );
}

function subcreator_extract_text_from_mogrt_item(item) {
  // // Read MOGRT subtitle text from real text-bearing controls and metadata only.
  if (!item) {
    return "";
  }

  function rememberReadableText(value) {
    var normalizedValue = subcreator_trim_string(String(value || "").replace(/\r/g, "\n"));
    if (!normalizedValue) {
      return "";
    }
    if (subcreator_is_default_caption_label(normalizedValue) || subcreator_is_non_textual_mogrt_label(normalizedValue)) {
      return "";
    }
    return normalizedValue;
  }

  var componentText = rememberReadableText(subcreator_extract_text_from_item_components(item));
  if (componentText) {
    return componentText;
  }

  var methodNames = ["getSourceText", "getText", "getFormattedText"];
  for (var methodIndex = 0; methodIndex < methodNames.length; methodIndex += 1) {
    var methodName = methodNames[methodIndex];
    try {
      if (typeof item[methodName] === "function") {
        var methodText = rememberReadableText(item[methodName]());
        if (methodText) {
          return methodText;
        }
      }
    } catch (methodError) {}
  }

  var propNames = ["captionText", "sourceText", "subtitleText", "text", "value"];
  for (var propIndex = 0; propIndex < propNames.length; propIndex += 1) {
    var propName = propNames[propIndex];
    try {
      if (typeof item[propName] !== "undefined") {
        var propText = rememberReadableText(item[propName]);
        if (propText) {
          return propText;
        }
      }
    } catch (propError) {}
  }

  try {
    if (item.projectItem && typeof item.projectItem.getProjectMetadata === "function") {
      var metadata = String(item.projectItem.getProjectMetadata() || "");
      var metadataText = rememberReadableText(subcreator_extract_text_from_metadata_blob(metadata));
      if (metadataText) {
        return metadataText;
      }
    }
  } catch (metadataError) {}

  try {
    if (item.projectItem && typeof item.projectItem.getXMPMetadata === "function") {
      var xmpMetadata = String(item.projectItem.getXMPMetadata() || "");
      var xmpText = rememberReadableText(subcreator_extract_text_from_metadata_blob(xmpMetadata));
      if (xmpText) {
        return xmpText;
      }
    }
  } catch (xmpMetadataError) {}

  return "";
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
    key.indexOf("text layer") !== -1 ||
    key.indexOf("textlayer") !== -1 ||
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

function subcreator_normalize_text_for_compare(textValue) {
  // // Compare text readback safely across Premiere payload variants that may use LF, CR, or extra whitespace.
  return subcreator_trim_string(String(textValue || "").replace(/\r/g, "\n").replace(/\s+/g, " "));
}

function subcreator_debug_push_limited(debugLines, line, maxLines) {
  // // Keep host debug output compact while still exposing the first useful diagnostics for problematic templates.
  if (!debugLines || typeof debugLines.push !== "function") {
    return;
  }

  var limit = Number(maxLines);
  if (isNaN(limit) || limit < 1) {
    limit = 80;
  }

  if (debugLines.length >= limit) {
    return;
  }

  debugLines.push(String(line || ""));
}

function subcreator_debug_scan_text_like_properties(propertyCollection, debugLines, debugPrefix, maxItems) {
  // // Inspect nearby text-like properties so Premiere-authored templates can reveal which control actually drives the visible text.
  if (!propertyCollection || typeof propertyCollection.numItems !== "number") {
    return;
  }

  var limit = Number(maxItems);
  if (isNaN(limit) || limit < 1) {
    limit = 12;
  }

  var scanned = 0;

  function walk(collection, pathPrefix) {
    if (!collection || typeof collection.numItems !== "number" || scanned >= limit) {
      return;
    }

    for (var propertyIndex = 0; propertyIndex < collection.numItems; propertyIndex += 1) {
      if (scanned >= limit) {
        return;
      }

      var property = collection[propertyIndex];
      if (!property) {
        continue;
      }

      var path = pathPrefix ? pathPrefix + "." + String(propertyIndex) : String(propertyIndex);
      var displayName = subcreator_trim_string(String(property.displayName || ""));
      var key = displayName.toLowerCase();
      var rawValue = "";
      var hasValue = false;

      if (typeof property.getValue === "function") {
        try {
          rawValue = property.getValue();
          hasValue = true;
        } catch (readError) {
          hasValue = false;
        }
      }

      var looksInteresting =
        key.indexOf("text") !== -1 ||
        key.indexOf("source") !== -1 ||
        key.indexOf("caption") !== -1 ||
        key.indexOf("layer") !== -1 ||
        (hasValue && typeof rawValue === "string" && String(rawValue || "").length > 0 && String(rawValue).length < 80);

      if (looksInteresting) {
        subcreator_debug_push_limited(
          debugLines,
          String(debugPrefix || "") +
            " candidate path=" +
            path +
            " name=" +
            String(displayName || "<unnamed>") +
            " rawType=" +
            typeof rawValue +
            " raw=" +
            subcreator_visual_preview_debug_value(rawValue, 120),
          120
        );
        scanned += 1;
      }

      if (property.properties && typeof property.properties.numItems === "number" && property.properties.numItems > 0) {
        walk(property.properties, path);
      }
    }
  }

  walk(propertyCollection, "");
}

function subcreator_property_readback_matches_text(property, expectedText) {
  // // Validate text writes by reading the property back, so silent setValue failures do not count as success.
  if (!property || typeof property.getValue !== "function") {
    return false;
  }

  var expected = subcreator_normalize_text_for_compare(expectedText);
  if (!expected) {
    return false;
  }

  try {
    var readbackValue = property.getValue();
    var extracted = subcreator_extract_text_from_property_value(readbackValue);
    var normalizedExtracted = subcreator_normalize_text_for_compare(extracted || readbackValue);
    return normalizedExtracted === expected;
  } catch (readbackError) {}

  return false;
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

function subcreator_try_apply_mogrt_text_property_raw_value(property, sourceRawValue, textValue, debugLines, debugPrefix) {
  // // Apply text using a provided source payload so text-document styles can be preserved during clip rebuilds.
  var displayName = property.displayName || "";
  var rawValue = sourceRawValue;

  if (!subcreator_should_try_text_property(displayName, rawValue)) {
    return false;
  }

  var plainTextString = String(textValue || "");
  var textString = subcreator_normalize_caption_text(textValue);
  var beforePreview = subcreator_visual_preview_debug_value(rawValue, 140);
  var prefix = subcreator_trim_string(String(debugPrefix || ""));
  var binaryPayloadCandidate = typeof rawValue === "string" ? subcreator_find_binary_text_payload_candidate(rawValue) : null;

  function debugAttempt(label, applied) {
    // // Capture text property readback to distinguish real text writes from writes that only touch a hidden control.
    var afterPreview = "";
    try {
      afterPreview = subcreator_visual_preview_debug_value(property.getValue(), 140);
    } catch (readbackError) {
      afterPreview = "<readback unavailable>";
    }

    subcreator_debug_push_limited(
      debugLines,
      prefix +
        " textprop name=" +
        String(displayName) +
        " mode=" +
        String(label) +
        " applied=" +
        (applied ? "true" : "false") +
        " rawType=" +
        typeof rawValue +
        " before=" +
        beforePreview +
        " after=" +
        afterPreview,
      120
    );
  }

  if (binaryPayloadCandidate) {
    try {
      var patchedBinaryPayload = subcreator_try_patch_binary_text_payload(rawValue, textValue);
      if (patchedBinaryPayload) {
        property.setValue(patchedBinaryPayload, true);
        if (subcreator_property_readback_matches_text(property, plainTextString)) {
          debugAttempt("binary_payload", true);
          return true;
        }
        // // Premiere can accept this binary write but collapse Source Text to an unreadable token; do not count that as a text update.
        debugAttempt("binary_payload_no_readback", false);
        return false;
      }
    } catch (binaryPayloadError) {}
  }

  if (typeof rawValue === "string" && rawValue.indexOf("{") === -1) {
    // // Premiere-authored MOGRT text controls often behave like plain string parameters and should keep their existing style untouched.
    if (!binaryPayloadCandidate) {
      var plainTextCandidates = [
        { value: textString, useRefresh: true },
        { value: textString, useRefresh: false },
        { value: plainTextString, useRefresh: true },
        { value: plainTextString, useRefresh: false }
      ];

      for (var candidateIndex = 0; candidateIndex < plainTextCandidates.length; candidateIndex += 1) {
        var textCandidate = plainTextCandidates[candidateIndex];
        try {
          property.setValue(textCandidate.value, textCandidate.useRefresh);
          if (subcreator_property_readback_matches_text(property, plainTextString)) {
            debugAttempt("plain_string_" + String(candidateIndex), true);
            return true;
          }
        } catch (plainSetError) {}
      }
    } else {
      debugAttempt("binary_payload_failed", false);
      return false;
    }
  }

  if (rawValue && typeof rawValue === "object") {
    try {
      var objectCopy = JSON.parse(JSON.stringify(rawValue));
      if (subcreator_try_set_json_text_payload(objectCopy, textString)) {
        property.setValue(objectCopy, true);
        if (subcreator_property_readback_matches_text(property, plainTextString)) {
          debugAttempt("object_copy", true);
          return true;
        }
      }
    } catch (objectJsonError) {}

    try {
      if (subcreator_try_set_json_text_payload(rawValue, textString)) {
        property.setValue(rawValue, true);
        if (subcreator_property_readback_matches_text(property, plainTextString)) {
          debugAttempt("object_direct", true);
          return true;
        }
      }
    } catch (objectDirectError) {}
  }

  if (typeof rawValue === "string" && rawValue.indexOf("{") !== -1) {
    try {
      var parsed = JSON.parse(rawValue);
      if (subcreator_try_set_json_text_payload(parsed, textString)) {
        property.setValue(JSON.stringify(parsed), true);
        if (subcreator_property_readback_matches_text(property, plainTextString)) {
          debugAttempt("json_string", true);
          return true;
        }
      }
    } catch (jsonError) {}

    try {
      var patchedRaw = subcreator_try_patch_text_json_string(rawValue, textString);
      if (patchedRaw) {
        property.setValue(patchedRaw, true);
        if (subcreator_property_readback_matches_text(property, plainTextString)) {
          debugAttempt("json_patch", true);
          return true;
        }
      }
    } catch (patchError) {}
  }

  try {
    property.setValue(plainTextString, true);
    if (subcreator_property_readback_matches_text(property, plainTextString)) {
      debugAttempt("fallback_plain", true);
      return true;
    }
  } catch (setError) {}

  debugAttempt("failed", false);

  return false;
}

function subcreator_try_set_mogrt_text_property(property, textValue, templatePayloadState, debugLines, debugPrefix) {
  // // Apply text to a property, supporting strings, JSON strings, and object payloads.
  var rawValue = "";

  if (typeof property.getValue === "function") {
    try {
      rawValue = property.getValue();
    } catch (getError) {
      rawValue = "";
    }
  }

  var templateRawValue = subcreator_resolve_template_text_payload_raw_value(templatePayloadState, property.displayName || "");
  var sourceRawValue = templateRawValue || rawValue;
  return subcreator_try_apply_mogrt_text_property_raw_value(property, sourceRawValue, textValue, debugLines, debugPrefix);
}

function subcreator_sleep_ms(milliseconds) {
  // // Give Premiere time to finish initializing newly imported MOGRT controls before writing text into them.
  try {
    if (typeof $ !== "undefined" && $ && typeof $.sleep === "function") {
      $.sleep(Math.max(0, Number(milliseconds) || 0));
    }
  } catch (sleepError) {}
}

function subcreator_collect_unique_mogrt_components_with_retry(trackItem, debugLines, debugPrefix) {
  // // Premiere-authored MOGRTs can expose their component tree a few frames after importMGT returns.
  var components = [];
  for (var attemptIndex = 0; attemptIndex < 6; attemptIndex += 1) {
    components = subcreator_collect_unique_mogrt_components(trackItem);
    if (components.length > 0) {
      if (attemptIndex > 0) {
        subcreator_debug_push_limited(
          debugLines,
          String(debugPrefix || "") + " components_ready_after_retry=" + String(attemptIndex),
          120
        );
      }
      return components;
    }
    subcreator_sleep_ms(80);
  }

  return components;
}

function subcreator_mogrt_has_nonempty_text_property(trackItem, debugLines, debugPrefix) {
  // // Validate pre-baked Premiere MOGRT imports by requiring a real non-empty text control after import.
  var components = subcreator_collect_unique_mogrt_components_with_retry(trackItem, debugLines, debugPrefix);

  function walk(collection) {
    if (!collection || typeof collection.numItems !== "number") {
      return false;
    }

    for (var propertyIndex = 0; propertyIndex < collection.numItems; propertyIndex += 1) {
      var property = collection[propertyIndex];
      if (!property) {
        continue;
      }

      var rawValue = "";
      var hasRawValue = false;
      if (typeof property.getValue === "function") {
        try {
          rawValue = property.getValue();
          hasRawValue = true;
        } catch (readError) {
          hasRawValue = false;
        }
      }

      if (
        hasRawValue &&
        subcreator_should_try_text_property(property.displayName || "", rawValue) &&
        subcreator_trim_string(String(rawValue || "")).length > 0
      ) {
        return true;
      }

      if (property.properties && typeof property.properties.numItems === "number" && property.properties.numItems > 0 && walk(property.properties)) {
        return true;
      }
    }

    return false;
  }

  for (var componentIndex = 0; componentIndex < components.length; componentIndex += 1) {
    var component = components[componentIndex];
    if (component && component.properties && walk(component.properties)) {
      return true;
    }
  }

  subcreator_debug_push_limited(debugLines, String(debugPrefix || "") + " baked_text_validation_failed", 120);
  return false;
}

function subcreator_collect_mogrt_text_payload_snapshots_recursive(propertyCollection, pathPrefix, outList) {
  // // Capture text-bearing raw payloads so rebuilt MOGRT clips can reuse the original text style document state.
  if (!propertyCollection || typeof propertyCollection.numItems !== "number") {
    return;
  }

  for (var index = 0; index < propertyCollection.numItems; index += 1) {
    var property = propertyCollection[index];
    if (!property) {
      continue;
    }

    var currentPath = pathPrefix ? pathPrefix + "." + String(index) : String(index);
    var rawValue = undefined;
    var hasValue = false;

    if (typeof property.getValue === "function") {
      try {
        rawValue = property.getValue();
        hasValue = true;
      } catch (readError) {
        hasValue = false;
      }
    }

    if (hasValue && subcreator_should_try_text_property(property.displayName || "", rawValue)) {
      outList.push({
        path: currentPath,
        rawValue: rawValue
      });
    }

    if (property.properties && typeof property.properties.numItems === "number" && property.properties.numItems > 0) {
      subcreator_collect_mogrt_text_payload_snapshots_recursive(property.properties, currentPath, outList);
    }
  }
}

function subcreator_build_text_clone_changes_from_track_item(trackItem) {
  // // Snapshot original text-document payloads from one subtitle clip so font/style survives text rebuilds.
  var component = subcreator_get_mogrt_component_from_track_item(trackItem);
  if (!component || !component.properties) {
    return [];
  }

  var snapshots = [];
  subcreator_collect_mogrt_text_payload_snapshots_recursive(component.properties, "", snapshots);
  return snapshots;
}

function subcreator_apply_text_clone_changes_to_track_item(trackItem, changes, textValue, debugLines) {
  // // Reapply captured text-document payloads with new text content to preserve style on rebuilt MOGRT clips.
  if (!trackItem || !changes || typeof changes.length !== "number") {
    return {
      updatedCount: 0,
      failedCount: 0
    };
  }

  var component = subcreator_get_mogrt_component_from_track_item(trackItem);
  if (!component || !component.properties) {
    return {
      updatedCount: 0,
      failedCount: changes.length
    };
  }

  var updatedCount = 0;
  var failedCount = 0;

  for (var changeIndex = 0; changeIndex < changes.length; changeIndex += 1) {
    var change = changes[changeIndex] || {};
    var path = subcreator_trim_string(String(change.path || ""));
    if (!path) {
      failedCount += 1;
      continue;
    }

    var property = subcreator_find_property_by_path(component.properties, path);
    if (!property || typeof property.setValue !== "function") {
      failedCount += 1;
      continue;
    }

    var applied = false;
    try {
      applied = subcreator_try_apply_mogrt_text_property_raw_value(property, change.rawValue, textValue, debugLines, "text_clone path=" + path);
    } catch (textCloneError) {
      applied = false;
    }

    if (applied) {
      updatedCount += 1;
    } else {
      if (debugLines && typeof debugLines.push === "function") {
        debugLines.push("text tab text clone failed path=" + path);
      }
      failedCount += 1;
    }
  }

  return {
    updatedCount: updatedCount,
    failedCount: failedCount
  };
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

    // // AE-authored subtitle templates commonly expose this dropdown as `0=Words / 1=Lines`, even when other menus look 1-based.
    // // Prefer the 0-based mapping first so generated clips and exports keep the intended highlight animation state.
    var currentValue = NaN;
    try {
      currentValue = Number(property.getValue());
    } catch (readError) {}

    var candidateValues = [];
    function rememberCandidate(value) {
      // // Keep fallbacks unique so we can try both enum bases without repeating the same write.
      for (var candidateIndex = 0; candidateIndex < candidateValues.length; candidateIndex += 1) {
        if (candidateValues[candidateIndex] === value) {
          return;
        }
      }
      candidateValues.push(value);
    }

    var preferZeroBased = true;
    if (!isNaN(currentValue)) {
      if (currentValue === 2) {
        preferZeroBased = false;
      } else if (currentValue === 0 || currentValue === 1) {
        preferZeroBased = true;
      }
    }

    if (mode === "word") {
      if (preferZeroBased) {
        rememberCandidate(0);
        rememberCandidate(1);
      } else {
        rememberCandidate(1);
        rememberCandidate(0);
      }
    } else {
      if (preferZeroBased) {
        rememberCandidate(1);
        rememberCandidate(2);
      } else {
        rememberCandidate(2);
        rememberCandidate(1);
      }
    }

    for (var highlightIndex = 0; highlightIndex < candidateValues.length; highlightIndex += 1) {
      try {
        var highlightValue = candidateValues[highlightIndex];
        property.setValue(highlightValue, true);
        try {
          if (typeof property.getValue === "function" && Number(property.getValue()) !== highlightValue) {
            continue;
          }
        } catch (readbackError) {}
        return true;
      } catch (highlightError) {}
    }
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
  // // Apply layout controls (characters/lines/clip duration) when available in a template.
  if (!property || typeof property.setValue !== "function" || !styleConfig) {
    return false;
  }

  var key = String(property.displayName || "").toLowerCase();
  var maxChars = Number(styleConfig.maxCharsPerLine || 0);
  var maxLines = Number(styleConfig.linesPerCaption || 0);
  var clipDurationSeconds = Number(styleConfig.clipDurationSeconds || 0);

  if (
    !isNaN(clipDurationSeconds) &&
    clipDurationSeconds > 0 &&
    (key === "clip duration" ||
      key.indexOf("clip duration") !== -1 ||
      key.indexOf("caption duration") !== -1 ||
      key.indexOf("subtitle duration") !== -1 ||
      key.indexOf("duree clip") !== -1)
  ) {
    try {
      // // Feed opt-in AE test templates with the real subtitle clip duration instead of forcing them to infer it from MOGRT comp timing.
      property.setValue(clipDurationSeconds, true);
      return true;
    } catch (clipDurationError) {}
  }

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

  return false;
}

function subcreator_clone_style_config_with_clip_duration(styleConfig, clipDurationSeconds) {
  // // Keep existing template style options intact while adding a cue-specific clip duration override for opt-in AE MOGRTs.
  var cloned = {};
  if (styleConfig && typeof styleConfig === "object") {
    for (var styleKey in styleConfig) {
      if (!styleConfig.hasOwnProperty(styleKey)) {
        continue;
      }
      cloned[styleKey] = styleConfig[styleKey];
    }
  }

  var safeDuration = Number(clipDurationSeconds);
  if (!isNaN(safeDuration) && safeDuration > 0) {
    cloned.clipDurationSeconds = safeDuration;
  }

  return cloned;
}

function subcreator_try_set_controls_recursively(
  propertyCollection,
  textValue,
  animationMode,
  styleConfig,
  templatePayloadState,
  skipTextUpdates,
  stats,
  debugLines,
  debugPrefix,
  passMode
) {
  // // Traverse nested Essential Graphics property groups in two passes so text is written last and can persist dependent animation metadata.
  if (!propertyCollection || typeof propertyCollection.numItems !== "number") {
    return;
  }

  var normalizedPassMode = subcreator_trim_string(String(passMode || "")).toLowerCase();

  for (var i = 0; i < propertyCollection.numItems; i += 1) {
    var property = propertyCollection[i];
    if (!property) {
      continue;
    }

    if (typeof property.setValue === "function") {
      if (normalizedPassMode !== "text") {
        if (subcreator_try_set_animation_mode_property(property, animationMode)) {
          stats.animationUpdates += 1;
        }

        if (subcreator_try_set_layout_property(property, styleConfig)) {
          stats.layoutUpdates += 1;
        }
      }

      if (
        normalizedPassMode !== "non_text" &&
        !skipTextUpdates &&
        subcreator_try_set_mogrt_text_property(property, textValue, templatePayloadState, debugLines, debugPrefix)
      ) {
        stats.textUpdates += 1;
      }
    }

    if (property.properties && typeof property.properties.numItems === "number" && property.properties.numItems > 0) {
      subcreator_try_set_controls_recursively(
        property.properties,
        textValue,
        animationMode,
        styleConfig,
        templatePayloadState,
        skipTextUpdates,
        stats,
        debugLines,
        debugPrefix,
        normalizedPassMode
      );
    }
  }
}

function subcreator_try_set_mogrt_controls(trackItem, textValue, animationMode, styleConfig, templateTextPayloads, skipTextUpdates, debugLines, debugPrefix) {
  // // Update text + animation related controls on inserted MOGRT components.
  if (!trackItem || !textValue) {
    return {
      textUpdates: 0,
      animationUpdates: 0,
      layoutUpdates: 0
    };
  }

  var components = subcreator_collect_unique_mogrt_components_with_retry(trackItem, debugLines, debugPrefix);
  if (components.length < 1) {
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
  var templatePayloadState = subcreator_create_template_text_payload_state(templateTextPayloads);

  for (var componentIndex = 0; componentIndex < components.length; componentIndex += 1) {
    var component = components[componentIndex];
    if (!component || !component.properties || component.properties.numItems < 1) {
      continue;
    }

    if (debugLines && typeof debugLines.push === "function") {
      subcreator_debug_scan_text_like_properties(
        component.properties,
        debugLines,
        debugPrefix ? debugPrefix + " component=" + String(componentIndex) : "component=" + String(componentIndex),
        10
      );
    }

    subcreator_try_set_controls_recursively(
      component.properties,
      textValue,
      animationMode,
      styleConfig,
      templatePayloadState,
      Boolean(skipTextUpdates),
      stats,
      debugLines,
      debugPrefix ? debugPrefix + " component=" + String(componentIndex) : "component=" + String(componentIndex),
      "non_text"
    );
  }

  for (var textComponentIndex = 0; textComponentIndex < components.length; textComponentIndex += 1) {
    var textComponent = components[textComponentIndex];
    if (!textComponent || !textComponent.properties || textComponent.properties.numItems < 1) {
      continue;
    }

    subcreator_try_set_controls_recursively(
      textComponent.properties,
      textValue,
      animationMode,
      styleConfig,
      templatePayloadState,
      Boolean(skipTextUpdates),
      stats,
      debugLines,
      debugPrefix ? debugPrefix + " component=" + String(textComponentIndex) : "component=" + String(textComponentIndex),
      "text"
    );
  }

  if (!skipTextUpdates && stats.textUpdates < 1) {
    for (var retryIndex = 0; retryIndex < 3; retryIndex += 1) {
      subcreator_sleep_ms(120);
      components = subcreator_collect_unique_mogrt_components(trackItem);
      for (var retryComponentIndex = 0; retryComponentIndex < components.length; retryComponentIndex += 1) {
        var retryComponent = components[retryComponentIndex];
        if (!retryComponent || !retryComponent.properties || retryComponent.properties.numItems < 1) {
          continue;
        }

        subcreator_try_set_controls_recursively(
          retryComponent.properties,
          textValue,
          animationMode,
          styleConfig,
          templatePayloadState,
          false,
          stats,
          debugLines,
          debugPrefix ? debugPrefix + " text_retry=" + String(retryIndex + 1) : "text_retry=" + String(retryIndex + 1),
          "text"
        );
      }

      if (stats.textUpdates > 0) {
        subcreator_debug_push_limited(
          debugLines,
          String(debugPrefix || "") + " text_ready_after_retry=" + String(retryIndex + 1),
          120
        );
        break;
      }
    }
  }

  subcreator_debug_push_limited(
    debugLines,
    String(debugPrefix || "") +
      " extracted_text=" +
      subcreator_visual_preview_debug_value(subcreator_extract_text_from_mogrt_item(trackItem), 180),
    120
  );

  return stats;
}

function subcreator_build_visual_clone_changes_from_track_item(trackItem) {
  // Snapshot the same filtered and deduplicated visual controls as the visual editor so AE duplicate components do not reapply hidden defaults.
  var rawComponents = subcreator_get_mogrt_components_from_track_item(trackItem);
  if (rawComponents.length < 1) {
    return [];
  }

  var properties = [];
  var seenComponentSignatures = {};
  subcreator_visual_reset_group_sequence_axis_preferences();
  subcreator_visual_reset_text_style_option_cache();
  for (var componentIndex = 0; componentIndex < rawComponents.length; componentIndex += 1) {
    var componentGroupPath = rawComponents.length > 1 ? subcreator_visual_get_component_group_label(rawComponents[componentIndex], componentIndex) : "";
    var componentProperties = [];
    subcreator_collect_mogrt_visual_properties_recursive(
      rawComponents[componentIndex] ? rawComponents[componentIndex].properties : null,
      subcreator_visual_build_component_path_prefix(componentIndex),
      componentGroupPath,
      componentProperties
    );

    componentProperties = subcreator_visual_filter_property_descriptors(componentProperties);
    var componentSignature = subcreator_visual_build_component_descriptor_signature(componentProperties);
    if (componentSignature && typeof seenComponentSignatures[componentSignature] !== "undefined") {
      continue;
    }

    seenComponentSignatures[componentSignature] = componentIndex;
    for (var componentPropertyIndex = 0; componentPropertyIndex < componentProperties.length; componentPropertyIndex += 1) {
      properties.push(componentProperties[componentPropertyIndex]);
    }
  }

  var changes = [];
  for (var propertyIndex = 0; propertyIndex < properties.length; propertyIndex += 1) {
    var property = properties[propertyIndex];
    if (!property || !property.path) {
      continue;
    }

    changes.push({
      path: property.path,
      valueType: property.valueType,
      controlKind: property.controlKind,
      value: property.value,
      vectorScale: property.vectorScale || null,
      vectorMode: property.vectorMode || null
    });
  }

  return changes;
}

function subcreator_apply_visual_changes_to_track_item(trackItem, changes, debugLines) {
  // // Reuse visual-editor setters so rebuilt text clips inherit the source clip appearance as closely as possible.
  if (!trackItem || !changes || typeof changes.length !== "number") {
    return {
      updatedCount: 0,
      failedCount: 0
    };
  }

  var components = subcreator_get_mogrt_components_from_track_item(trackItem);
  if (components.length < 1) {
    return {
      updatedCount: 0,
      failedCount: changes.length
    };
  }

  var updatedCount = 0;
  var failedCount = 0;

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

    var resolvedProperty = subcreator_visual_resolve_property_from_track_item(trackItem, resolvedPath);
    var property = resolvedProperty ? resolvedProperty.property : null;
    if (!property || typeof property.setValue !== "function") {
      failedCount += 1;
      continue;
    }

    var applied = false;
    if (virtualTextStyleTarget) {
      try {
        applied = subcreator_try_set_mogrt_text_style_property(property, virtualTextStyleTarget.styleKey, value, {});
      } catch (textStyleError) {}
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
          applied = true;
        }
      } catch (vectorError) {}
    } else if ((controlKind === "slider" || controlKind === "number") && valueType === "number") {
      try {
        applied = subcreator_visual_try_set_numeric_property(
          property,
          Number(value),
          debugLines,
          path + " clone"
        );
      } catch (numericError) {}
    }

    if (!applied && controlKind !== "color" && !virtualTextStyleTarget) {
      try {
        var normalizedValue = subcreator_normalize_visual_payload_value(valueType, value);
        property.setValue(normalizedValue, true);
        applied = true;
      } catch (setError) {
        applied = false;
      }
    }

    if (applied) {
      updatedCount += 1;
    } else {
      if (debugLines && typeof debugLines.push === "function") {
        debugLines.push("text tab visual clone failed path=" + path + " kind=" + controlKind);
      }
      failedCount += 1;
    }
  }

  return {
    updatedCount: updatedCount,
    failedCount: failedCount
  };
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

function subcreator_get_sequence_frame_duration_seconds(sequence) {
  // // Read Premiere's per-frame Time value so timing tolerances follow the active sequence frame rate.
  try {
    var settings = sequence && typeof sequence.getSettings === "function" ? sequence.getSettings() : null;
    var frameDuration = subcreator_to_seconds(settings && settings.videoFrameRate);
    if (!isNaN(frameDuration) && frameDuration >= 1 / 240 && frameDuration <= 1) {
      return frameDuration;
    }
  } catch (frameDurationError) {}

  return 1 / 30;
}

function subcreator_snap_nearby_caption_end(endSeconds, nextStartSeconds, frameDurationSeconds) {
  // // Close only sub-frame/one-frame gaps that QE razor rounding would otherwise make visibly empty.
  var safeEnd = Number(endSeconds);
  var safeNextStart = Number(nextStartSeconds);
  var safeFrameDuration = Number(frameDurationSeconds);
  if (isNaN(safeEnd) || isNaN(safeNextStart) || isNaN(safeFrameDuration) || safeFrameDuration <= 0) {
    return safeEnd;
  }

  var gapSeconds = safeNextStart - safeEnd;
  if (gapSeconds >= 0 && gapSeconds <= safeFrameDuration + 0.0005) {
    return safeNextStart;
  }

  return safeEnd;
}

function subcreator_snap_seconds_to_nearest_frame(seconds, frameDurationSeconds) {
  // // Use one shared frame boundary for importMGT placement and QE razor cuts instead of their opposite sub-frame rounding.
  var safeSeconds = Number(seconds);
  var safeFrameDuration = Number(frameDurationSeconds);
  if (isNaN(safeSeconds) || isNaN(safeFrameDuration) || safeFrameDuration <= 0) {
    return safeSeconds;
  }

  return Math.round(safeSeconds / safeFrameDuration) * safeFrameDuration;
}

function subcreator_track_item_starts_near_seconds(trackItem, startSeconds, toleranceSeconds) {
  // // Validate imported clips against the requested insertion time so importMGT mismatches do not get accepted silently.
  var candidateStart = subcreator_to_seconds(trackItem && (trackItem.start || trackItem.inPoint || trackItem.startTime));
  var requestedStart = Number(startSeconds);
  var tolerance = Number(toleranceSeconds);
  if (isNaN(candidateStart) || isNaN(requestedStart)) {
    return false;
  }

  if (isNaN(tolerance) || tolerance <= 0) {
    tolerance = 0.2;
  }

  return Math.abs(candidateStart - requestedStart) <= tolerance;
}

function subcreator_track_item_ends_near_seconds(trackItem, endSeconds, toleranceSeconds) {
  // // Prefer timeline-end trimming over source outPoint edits so Premiere-authored MOGRT keyframes stay intact.
  var candidateEnd = subcreator_to_seconds(trackItem && (trackItem.end || trackItem.endTime));
  var requestedEnd = Number(endSeconds);
  var tolerance = Number(toleranceSeconds);
  if (isNaN(candidateEnd) || isNaN(requestedEnd)) {
    return false;
  }

  if (isNaN(tolerance) || tolerance <= 0) {
    tolerance = 0.2;
  }

  return Math.abs(candidateEnd - requestedEnd) <= tolerance;
}

function subcreator_find_track_clip_near_start(track, startSeconds, projectItem) {
  // // Resolve a timeline fragment by start time after QE razor invalidates the original TrackItem reference.
  var clips = subcreator_collection_to_array(track ? track.clips : null);
  var safeStart = Number(startSeconds);
  var bestItem = null;
  var bestDistance = Number.POSITIVE_INFINITY;

  for (var clipIndex = 0; clipIndex < clips.length; clipIndex += 1) {
    var candidate = clips[clipIndex];
    if (!candidate) {
      continue;
    }
    if (projectItem && candidate.projectItem && candidate.projectItem !== projectItem) {
      var sourceNodeId = subcreator_trim_string(String(projectItem.nodeId || ""));
      var candidateNodeId = subcreator_trim_string(String(candidate.projectItem.nodeId || ""));
      if (!sourceNodeId || !candidateNodeId || sourceNodeId !== candidateNodeId) {
        continue;
      }
    }

    var candidateStart = subcreator_to_seconds(candidate.start || candidate.inPoint || candidate.startTime);
    var distance = Math.abs(candidateStart - safeStart);
    if (!isNaN(candidateStart) && distance <= 0.2 && distance < bestDistance) {
      bestItem = candidate;
      bestDistance = distance;
    }
  }

  return bestItem;
}

function subcreator_format_qe_razor_time(sequence, seconds) {
  // // Convert sequence seconds to the timecode string expected by QETrack.razor().
  try {
    var settings = sequence.getSettings();
    var razorTime = new Time();
    var frameDurationSeconds = subcreator_get_sequence_frame_duration_seconds(sequence);
    // // Stay just inside the chosen frame so floating-point conversion cannot format the preceding timecode frame.
    razorTime.seconds = Number(seconds) + frameDurationSeconds * 0.001;
    if (settings && settings.videoFrameRate && typeof razorTime.getFormatted === "function") {
      return razorTime.getFormatted(settings.videoFrameRate, settings.videoDisplayFormat);
    }
  } catch (formatRazorTimeError) {}

  return "";
}

function subcreator_try_razor_mogrt_duration(sequence, trackItem, videoTrackIndex, startSeconds, endSeconds, debugLines, debugPrefix) {
  // // Cut the imported MOGRT like a manual razor edit so its source duration stays available for later extension.
  var result = {
    applied: false,
    trackItem: trackItem,
    removedRightFragment: false,
    razorTimecode: ""
  };
  var safeStart = Number(startSeconds);
  var safeEnd = Number(endSeconds);
  var frameDurationSeconds = subcreator_get_sequence_frame_duration_seconds(sequence);
  var exactEndToleranceSeconds = Math.max(0.0005, Math.min(0.001, frameDurationSeconds * 0.1));
  var razorEndToleranceSeconds = Math.max(0.001, frameDurationSeconds * 0.51);
  var targetTrack = sequence && sequence.videoTracks ? sequence.videoTracks[videoTrackIndex] : null;
  var currentEnd = subcreator_to_seconds(trackItem && (trackItem.end || trackItem.endTime));

  if (!trackItem || !targetTrack || isNaN(safeStart) || isNaN(safeEnd) || safeEnd <= safeStart) {
    return result;
  }
  if (!isNaN(currentEnd) && Math.abs(currentEnd - safeEnd) <= exactEndToleranceSeconds) {
    result.applied = true;
    return result;
  }
  if (!isNaN(currentEnd) && currentEnd < safeEnd) {
    return result;
  }

  try {
    if (typeof app.enableQE === "function") {
      app.enableQE();
    }
    var razorTimecode = subcreator_format_qe_razor_time(sequence, safeEnd);
    result.razorTimecode = razorTimecode;
    if (
      !razorTimecode ||
      typeof qe === "undefined" ||
      !qe.project ||
      typeof qe.project.getActiveSequence !== "function"
    ) {
      return result;
    }

    var qeSequence = qe.project.getActiveSequence();
    var qeTrack = qeSequence && typeof qeSequence.getVideoTrackAt === "function"
      ? qeSequence.getVideoTrackAt(Number(videoTrackIndex))
      : null;
    if (!qeTrack || typeof qeTrack.razor !== "function") {
      return result;
    }

    var sourceProjectItem = trackItem.projectItem || null;
    qeTrack.razor(razorTimecode);

    var leftFragment = subcreator_find_track_clip_near_start(targetTrack, safeStart, sourceProjectItem);
    var rightFragment = subcreator_find_track_clip_near_start(targetTrack, safeEnd, sourceProjectItem);
    if (rightFragment && rightFragment !== leftFragment) {
      result.removedRightFragment = subcreator_remove_track_item_without_ripple(rightFragment);
    }
    var refreshedLeftFragment = subcreator_find_track_clip_near_start(targetTrack, safeStart, sourceProjectItem);
    if (refreshedLeftFragment) {
      leftFragment = refreshedLeftFragment;
      result.trackItem = refreshedLeftFragment;
    }

    result.applied =
      Boolean(leftFragment) &&
      result.removedRightFragment &&
      subcreator_track_item_ends_near_seconds(leftFragment, safeEnd, razorEndToleranceSeconds);
  } catch (razorDurationError) {}

  subcreator_debug_push_limited(
    debugLines,
    String(debugPrefix || "") +
      " duration_razor timecode=" +
      String(result.razorTimecode || "<unavailable>") +
      " right_removed=" +
      (result.removedRightFragment ? "true" : "false") +
      " applied=" +
      (result.applied ? "true" : "false"),
    120
  );
  return result;
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

function subcreator_try_import_mogrt(sequence, pathCandidates, startSeconds, videoTrackIndex, audioTrackIndex, debugLines, debugPrefix) {
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
  var targetTrack = sequence && sequence.videoTracks ? sequence.videoTracks[videoTrackIndex] : null;

  for (var pathIndex = 0; pathIndex < pathCandidates.length; pathIndex += 1) {
    var pathCandidate = pathCandidates[pathIndex];

    for (var timeIndex = 0; timeIndex < timeModes.length; timeIndex += 1) {
      var timeMode = timeModes[timeIndex];
      importResult.attempted += 1;
      var beforeItems = subcreator_collection_to_array(targetTrack ? targetTrack.clips : null);

      try {
        var insertedItem = sequence.importMGT(pathCandidate, timeMode.value, videoTrackIndex, audioTrackIndex);
        var resolvedItem = null;
        var insertedStart = subcreator_to_seconds(insertedItem && (insertedItem.start || insertedItem.inPoint || insertedItem.startTime));

        if (insertedItem && subcreator_track_item_starts_near_seconds(insertedItem, startSeconds, 0.2)) {
          resolvedItem = insertedItem;
        } else if (insertedItem && subcreator_try_set_mogrt_start(insertedItem, startSeconds)) {
          resolvedItem = insertedItem;
        } else if (targetTrack) {
          resolvedItem = subcreator_find_inserted_track_item(
            targetTrack,
            beforeItems,
            startSeconds,
            insertedItem && insertedItem.projectItem ? insertedItem.projectItem : null
          );
        }

        subcreator_debug_push_limited(
          debugLines,
          String(debugPrefix || "") +
            " import path=" +
            pathCandidate +
            " mode=" +
            String(timeMode.mode) +
            " requested=" +
            String(Number(startSeconds || 0)) +
            " insertedStart=" +
            String(insertedStart) +
            " resolvedStart=" +
            String(subcreator_to_seconds(resolvedItem && (resolvedItem.start || resolvedItem.inPoint || resolvedItem.startTime))) +
            " accepted=" +
            (resolvedItem && subcreator_track_item_starts_near_seconds(resolvedItem, startSeconds, 0.2) ? "true" : "false"),
          120
        );

        if (resolvedItem && subcreator_track_item_starts_near_seconds(resolvedItem, startSeconds, 0.2)) {
          importResult.trackItem = resolvedItem;
          importResult.usedPath = pathCandidate;
          importResult.usedTimeMode = timeMode.mode;
          return importResult;
        }

        if (insertedItem && !resolvedItem) {
          subcreator_remove_track_item_without_ripple(insertedItem);
        }
      } catch (importError) {}
    }
  }

  return importResult;
}

function subcreator_try_set_mogrt_start(trackItem, startSeconds) {
  // // Reposition imported MOGRT clips when Premiere returns them at the wrong insertion time for a given template.
  if (!trackItem) {
    return false;
  }

  var safeStart = Number(startSeconds);
  if (isNaN(safeStart)) {
    return false;
  }

  var applied = false;
  var startTime = null;
  try {
    startTime = new Time();
    startTime.seconds = safeStart;
  } catch (createTimeError) {
    startTime = null;
  }

  if (startTime) {
    try {
      trackItem.start = startTime;
      applied = true;
    } catch (startAssignError) {}

    try {
      if (trackItem.start && typeof trackItem.start.seconds !== "undefined") {
        trackItem.start.seconds = safeStart;
        applied = true;
      }
    } catch (startSecondsError) {}
  }

  try {
    if (typeof trackItem.move === "function") {
      trackItem.move(subcreator_seconds_to_ticks(safeStart));
      applied = true;
    }
  } catch (moveTicksError) {}

  try {
    if (typeof trackItem.move === "function") {
      trackItem.move(safeStart);
      applied = true;
    }
  } catch (moveSecondsError) {}

  return applied && subcreator_track_item_starts_near_seconds(trackItem, safeStart, 0.2);
}

function subcreator_remove_track_item_without_ripple(trackItem) {
  // // Delete one existing MOGRT clip while keeping later timeline items anchored in place.
  if (!trackItem || typeof trackItem.remove !== "function") {
    return false;
  }

  try {
    trackItem.remove(0, 0);
    return true;
  } catch (removeFullSignatureError) {}

  try {
    trackItem.remove(0);
    return true;
  } catch (removeOneArgError) {}

  try {
    trackItem.remove(false, false);
    return true;
  } catch (removeBoolSignatureError) {}

  try {
    trackItem.remove(false);
    return true;
  } catch (removeBoolOneArgError) {}

  try {
    trackItem.remove();
    return true;
  } catch (removeNoArgError) {}

  return false;
}

function subcreator_do_ranges_overlap(leftStart, leftEnd, rightStart, rightEnd) {
  // // Detect timeline overlap with a small tolerance so rebuilt MOGRTs do not fallback because of sub-frame rounding noise.
  var safeLeftStart = Number(leftStart);
  var safeLeftEnd = Number(leftEnd);
  var safeRightStart = Number(rightStart);
  var safeRightEnd = Number(rightEnd);
  var overlapToleranceSeconds = 0.01;
  if (isNaN(safeLeftStart) || isNaN(safeLeftEnd) || isNaN(safeRightStart) || isNaN(safeRightEnd)) {
    return false;
  }

  return safeLeftStart < safeRightEnd - overlapToleranceSeconds && safeLeftEnd > safeRightStart + overlapToleranceSeconds;
}

function subcreator_build_overlap_range(leftStart, leftEnd, rightStart, rightEnd) {
  // // Build the exact intersecting interval so overlap checks can distinguish one new collision from time already occupied by rebuilt clips.
  var overlapStart = Math.max(Number(leftStart), Number(rightStart));
  var overlapEnd = Math.min(Number(leftEnd), Number(rightEnd));
  if (!subcreator_do_ranges_overlap(leftStart, leftEnd, rightStart, rightEnd)) {
    return null;
  }
  return {
    startSeconds: overlapStart,
    endSeconds: overlapEnd
  };
}

function subcreator_is_range_fully_covered_by_track_items(rangeStart, rangeEnd, trackItems) {
  // // Treat overlap as safe when the whole conflicting interval was already occupied by the clips that will be removed and rebuilt.
  var coverageSegments = [];
  var safeRangeStart = Number(rangeStart);
  var safeRangeEnd = Number(rangeEnd);
  var coverageToleranceSeconds = 0.01;
  if (!(safeRangeEnd > safeRangeStart) || !trackItems || typeof trackItems.length !== "number") {
    return false;
  }

  for (var itemIndex = 0; itemIndex < trackItems.length; itemIndex += 1) {
    var trackItem = trackItems[itemIndex];
    if (!trackItem) {
      continue;
    }
    var itemStart = subcreator_to_seconds(trackItem.start || trackItem.inPoint || trackItem.startTime);
    var itemEnd = subcreator_to_seconds(trackItem.end || trackItem.outPoint || trackItem.endTime);
    var overlapRange = subcreator_build_overlap_range(itemStart, itemEnd, safeRangeStart, safeRangeEnd);
    if (!overlapRange) {
      continue;
    }
    coverageSegments.push(overlapRange);
  }

  if (coverageSegments.length < 1) {
    return false;
  }

  coverageSegments.sort(function (left, right) {
    return Number(left.startSeconds || 0) - Number(right.startSeconds || 0);
  });

  var coveredUntil = safeRangeStart;
  for (var segmentIndex = 0; segmentIndex < coverageSegments.length; segmentIndex += 1) {
    var segment = coverageSegments[segmentIndex];
    if (Number(segment.startSeconds || 0) > coveredUntil + coverageToleranceSeconds) {
      return false;
    }
    coveredUntil = Math.max(coveredUntil, Number(segment.endSeconds || 0));
    if (coveredUntil >= safeRangeEnd - coverageToleranceSeconds) {
      return true;
    }
  }

  return coveredUntil >= safeRangeEnd - coverageToleranceSeconds;
}

function subcreator_is_selected_track_item_reference(trackItem, selectedItems) {
  // // Keep reference-based selection matching isolated so overlap checks can safely skip the clips being rebuilt.
  for (var index = 0; index < selectedItems.length; index += 1) {
    if (selectedItems[index] === trackItem) {
      return true;
    }
  }
  return false;
}

function subcreator_find_text_rebuild_overlap(track, selectedItems, editedItems, replacedItems) {
  // // Abort risky text rebuilds only when they create one new overlap beyond the area already occupied by the clips being replaced.
  var trackItems = subcreator_collection_to_array(track ? track.clips : null);
  for (var trackIndex = 0; trackIndex < trackItems.length; trackIndex += 1) {
    var candidate = trackItems[trackIndex];
    if (!candidate || subcreator_is_selected_track_item_reference(candidate, selectedItems)) {
      continue;
    }

    var candidateStart = subcreator_to_seconds(candidate.start || candidate.inPoint || candidate.startTime);
    var candidateEnd = subcreator_to_seconds(candidate.end || candidate.outPoint || candidate.endTime);
    if (isNaN(candidateStart) || isNaN(candidateEnd) || candidateEnd <= candidateStart) {
      continue;
    }

    for (var editedIndex = 0; editedIndex < editedItems.length; editedIndex += 1) {
      var editedItem = editedItems[editedIndex] || {};
      var editedStart = Number(editedItem.startSeconds);
      var editedEnd = Number(editedItem.endSeconds);
      if (isNaN(editedStart) || isNaN(editedEnd) || editedEnd <= editedStart) {
        continue;
      }

      var overlapRange = subcreator_build_overlap_range(candidateStart, candidateEnd, editedStart, editedEnd);
      if (overlapRange) {
        if (subcreator_is_range_fully_covered_by_track_items(overlapRange.startSeconds, overlapRange.endSeconds, replacedItems)) {
          continue;
        }
        return {
          clipName: subcreator_trim_string(
            String((candidate.projectItem && candidate.projectItem.name) || candidate.name || "Clip")
          ),
          startSeconds: candidateStart,
          endSeconds: candidateEnd
        };
      }
    }
  }

  return null;
}

function subcreator_find_text_rebuild_following_clip(track, replacedItems, rangeEndSeconds) {
  // // Force fallback-track rebuild when any clip remains later on the source track, because importMGT lands at template default duration before the final trim is applied.
  var trackItems = subcreator_collection_to_array(track ? track.clips : null);
  var safeRangeEnd = Number(rangeEndSeconds);
  var overlapToleranceSeconds = 0.01;
  var closestCandidate = null;
  var closestStart = Number.POSITIVE_INFINITY;

  if (isNaN(safeRangeEnd)) {
    return null;
  }

  for (var trackIndex = 0; trackIndex < trackItems.length; trackIndex += 1) {
    var candidate = trackItems[trackIndex];
    if (!candidate || subcreator_is_selected_track_item_reference(candidate, replacedItems)) {
      continue;
    }

    var candidateStart = subcreator_to_seconds(candidate.start || candidate.inPoint || candidate.startTime);
    var candidateEnd = subcreator_to_seconds(candidate.end || candidate.outPoint || candidate.endTime);
    if (isNaN(candidateStart) || isNaN(candidateEnd) || candidateEnd <= candidateStart) {
      continue;
    }

    if (candidateEnd <= safeRangeEnd + overlapToleranceSeconds) {
      continue;
    }

    if (candidateStart < closestStart) {
      closestStart = candidateStart;
      closestCandidate = {
        clipName: subcreator_trim_string(
          String((candidate.projectItem && candidate.projectItem.name) || candidate.name || "Clip")
        ),
        startSeconds: candidateStart,
        endSeconds: candidateEnd
      };
    }
  }

  return closestCandidate;
}

function subcreator_find_empty_video_track_above(sequence, baseTrackIndex) {
  // // Find the first already-existing empty video track above the edited subtitles for safe rebuild fallback.
  if (!sequence || !sequence.videoTracks || isNaN(Number(baseTrackIndex))) {
    return -1;
  }

  for (var trackIndex = Number(baseTrackIndex) + 1; trackIndex < sequence.videoTracks.numTracks; trackIndex += 1) {
    var candidateTrack = sequence.videoTracks[trackIndex];
    var candidateItems = subcreator_collection_to_array(candidateTrack ? candidateTrack.clips : null);
    if (candidateItems.length < 1) {
      return trackIndex;
    }
  }

  return -1;
}

function subcreator_append_top_video_track_via_qe(sequence) {
  // // Append one new top video track through QE so subtitle rebuild fallback can avoid occupied tracks.
  var currentTracks = sequence && sequence.videoTracks ? Number(sequence.videoTracks.numTracks || 0) : 0;
  var inserted = false;

  try {
    if (typeof app.enableQE === "function") {
      app.enableQE();
      if (typeof qe !== "undefined" && qe.project && typeof qe.project.getActiveSequence === "function") {
        var qeSequence = qe.project.getActiveSequence();
        if (qeSequence && typeof qeSequence.addTracks === "function") {
          if (!inserted && currentTracks > 0) {
            try {
              // // Append one video track at the top so existing tracks keep their relative ordering.
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

          if (!inserted) {
            try {
              qeSequence.addTracks(1);
              inserted = true;
            } catch (signatureErrorSingleArg) {}
          }
        }
      }
    }
  } catch (error) {}

  var updatedTracks = sequence && sequence.videoTracks ? Number(sequence.videoTracks.numTracks || 0) : currentTracks;
  return {
    created: updatedTracks > currentTracks || inserted,
    beforeTracks: currentTracks,
    afterTracks: updatedTracks,
    index: updatedTracks > 0 ? updatedTracks - 1 : -1
  };
}

function subcreator_get_or_create_video_track_above_index(sequence, baseTrackIndex) {
  // // Prefer one empty track already above the edited subtitles, otherwise append a brand-new top track via QE.
  var emptyTrackAbove = subcreator_find_empty_video_track_above(sequence, baseTrackIndex);
  if (emptyTrackAbove >= 0) {
    return {
      index: emptyTrackAbove,
      created: false,
      beforeTracks: sequence && sequence.videoTracks ? Number(sequence.videoTracks.numTracks || 0) : 0,
      afterTracks: sequence && sequence.videoTracks ? Number(sequence.videoTracks.numTracks || 0) : 0
    };
  }

  return subcreator_append_top_video_track_via_qe(sequence);
}

function subcreator_find_inserted_track_item(track, beforeItems, startSeconds, projectItem) {
  // // Resolve the new clip object after overwrite/insert helpers that do not return a TrackItem reference.
  var afterItems = subcreator_collection_to_array(track ? track.clips : null);
  for (var afterIndex = 0; afterIndex < afterItems.length; afterIndex += 1) {
    var candidate = afterItems[afterIndex];
    var alreadyPresent = false;
    for (var beforeIndex = 0; beforeIndex < beforeItems.length; beforeIndex += 1) {
      if (beforeItems[beforeIndex] === candidate) {
        alreadyPresent = true;
        break;
      }
    }

    if (alreadyPresent) {
      continue;
    }

    var candidateStart = subcreator_to_seconds(candidate && (candidate.start || candidate.inPoint || candidate.startTime));
    if (Math.abs(candidateStart - Number(startSeconds || 0)) > 0.2) {
      continue;
    }

    if (projectItem && candidate && candidate.projectItem && candidate.projectItem !== projectItem) {
      continue;
    }

    return candidate;
  }

  var closestCandidate = null;
  var closestDistance = 999999;
  for (var candidateIndex = 0; candidateIndex < afterItems.length; candidateIndex += 1) {
    var fallbackCandidate = afterItems[candidateIndex];
    var fallbackStart = subcreator_to_seconds(
      fallbackCandidate && (fallbackCandidate.start || fallbackCandidate.inPoint || fallbackCandidate.startTime)
    );
    var fallbackDistance = Math.abs(fallbackStart - Number(startSeconds || 0));
    if (fallbackDistance < closestDistance) {
      closestDistance = fallbackDistance;
      closestCandidate = fallbackCandidate;
    }
  }

  return closestDistance <= 0.25 ? closestCandidate : null;
}

function subcreator_try_place_project_item_on_track(sequence, track, projectItem, startSeconds, videoTrackIndex, audioTrackIndex) {
  // // Try overwrite/insert variants so rebuilt clips can be recreated from the selected source ProjectItem.
  if (!track || !projectItem) {
    return null;
  }

  var beforeItems = subcreator_collection_to_array(track.clips);
  var timeObject = null;
  try {
    timeObject = new Time();
    timeObject.seconds = Number(startSeconds || 0);
  } catch (timeError) {
    timeObject = null;
  }
  var tickValue = subcreator_seconds_to_ticks(startSeconds);
  var rawSeconds = Number(startSeconds || 0);

  var attempts = [
    function () {
      if (typeof track.overwriteClip === "function") {
        track.overwriteClip(projectItem, tickValue);
        return true;
      }
      return false;
    },
    function () {
      if (typeof track.overwriteClip === "function") {
        track.overwriteClip(projectItem, rawSeconds);
        return true;
      }
      return false;
    },
    function () {
      if (typeof track.insertClip === "function") {
        track.insertClip(projectItem, tickValue, videoTrackIndex, audioTrackIndex);
        return true;
      }
      return false;
    },
    function () {
      if (typeof track.insertClip === "function") {
        track.insertClip(projectItem, rawSeconds, videoTrackIndex, audioTrackIndex);
        return true;
      }
      return false;
    },
    function () {
      if (typeof sequence.insertClip === "function") {
        sequence.insertClip(projectItem, tickValue, videoTrackIndex, audioTrackIndex);
        return true;
      }
      return false;
    },
    function () {
      if (typeof sequence.overwriteClip === "function") {
        sequence.overwriteClip(projectItem, tickValue, videoTrackIndex, audioTrackIndex);
        return true;
      }
      return false;
    },
    function () {
      if (timeObject && typeof track.insertClip === "function") {
        track.insertClip(projectItem, timeObject, videoTrackIndex, audioTrackIndex);
        return true;
      }
      return false;
    }
  ];

  for (var attemptIndex = 0; attemptIndex < attempts.length; attemptIndex += 1) {
    try {
      var attempted = attempts[attemptIndex]();
      if (!attempted) {
        continue;
      }
      var insertedItem = subcreator_find_inserted_track_item(track, beforeItems, startSeconds, projectItem);
      if (insertedItem) {
        return insertedItem;
      }
    } catch (insertError) {}
  }

  return null;
}

function subcreator_try_select_track_item(trackItem, deselectOthers) {
  // // Reselect rebuilt clips so the Text tab can continue operating on the updated subtitle block.
  if (!trackItem || typeof trackItem.setSelected !== "function") {
    return false;
  }

  try {
    trackItem.setSelected(true, deselectOthers === true);
    return true;
  } catch (selectError) {}

  try {
    trackItem.setSelected(1, deselectOthers === true ? 1 : 0);
    return true;
  } catch (selectNumericError) {}

  try {
    trackItem.setSelected(true);
    return true;
  } catch (selectSingleError) {}

  return false;
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

  var appendResult = subcreator_append_top_video_track_via_qe(sequence);
  var updatedTracks = appendResult.afterTracks;
  var created = appendResult.created;
  var highestEmptyAfter = subcreator_find_highest_empty_video_track_index(sequence.videoTracks);
  var fallbackTop = updatedTracks > 0 ? updatedTracks - 1 : 0;

  return {
    index: highestEmptyAfter >= 0 ? highestEmptyAfter : fallbackTop,
    created: created,
    beforeTracks: currentTracks,
    afterTracks: updatedTracks
  };
}

function subcreator_format_srt_timestamp(totalSeconds) {
  // // Convert seconds into SRT timecode for Premiere's native subtitle importer.
  var safeSeconds = Math.max(0, Number(totalSeconds) || 0);
  var totalMilliseconds = Math.round(safeSeconds * 1000);
  var hours = Math.floor(totalMilliseconds / 3600000);
  var minutes = Math.floor((totalMilliseconds % 3600000) / 60000);
  var seconds = Math.floor((totalMilliseconds % 60000) / 1000);
  var milliseconds = totalMilliseconds % 1000;
  return (
    String(hours).replace(/^(\d)$/, "0$1") +
    ":" +
    String(minutes).replace(/^(\d)$/, "0$1") +
    ":" +
    String(seconds).replace(/^(\d)$/, "0$1") +
    "," +
    ("00" + String(milliseconds)).slice(-3)
  );
}

function subcreator_normalize_srt_caption_text(value) {
  // // Keep caption text valid for SRT blocks while preserving planner-created line breaks.
  return subcreator_trim_string(String(value || ""))
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/\n\s*\n/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n[ \t]+/g, "\n");
}

function subcreator_normalize_native_subtitle_cues(cues) {
  // // Premiere's native SRT importer rejects overlapping cues, so normalize timeline timings before writing.
  var normalized = [];
  var source = [];
  for (var i = 0; i < cues.length; i += 1) {
    var cue = cues[i] || {};
    var startSeconds = Number(cue.startSeconds);
    var endSeconds = Number(cue.endSeconds);
    var text = subcreator_normalize_srt_caption_text(cue.text || "");
    if (!isFinite(startSeconds) || !isFinite(endSeconds) || endSeconds <= startSeconds || !text) {
      continue;
    }
    source.push({
      startSeconds: Math.max(0, startSeconds),
      endSeconds: Math.max(0, endSeconds),
      text: text
    });
  }

  source.sort(function (left, right) {
    return left.startSeconds - right.startSeconds || left.endSeconds - right.endSeconds;
  });

  var minimumDurationSeconds = 0.05;
  var gapSeconds = 0.001;
  for (var index = 0; index < source.length; index += 1) {
    var sourceCue = source[index];
    var normalizedStart = sourceCue.startSeconds;
    var normalizedEnd = Math.max(sourceCue.endSeconds, normalizedStart + minimumDurationSeconds);
    if (normalized.length > 0) {
      var previousCue = normalized[normalized.length - 1];
      var minimumStart = previousCue.endSeconds + gapSeconds;
      if (normalizedStart < minimumStart) {
        normalizedStart = minimumStart;
      }
      if (normalizedEnd <= normalizedStart) {
        normalizedEnd = normalizedStart + minimumDurationSeconds;
      }
    }

    normalized.push({
      startSeconds: normalizedStart,
      endSeconds: normalizedEnd,
      text: sourceCue.text
    });
  }

  return normalized;
}

function subcreator_serialize_cues_to_srt(cues) {
  // // Serialize generated caption cues to one SRT file for fast native subtitle track creation.
  var lines = [];
  var normalizedCues = subcreator_normalize_native_subtitle_cues(cues || []);
  for (var i = 0; i < normalizedCues.length; i += 1) {
    var cue = normalizedCues[i] || {};
    lines.push(
      String(lines.length + 1) +
        "\n" +
        subcreator_format_srt_timestamp(cue.startSeconds) +
        " --> " +
        subcreator_format_srt_timestamp(cue.endSeconds) +
        "\n" +
        cue.text
    );
  }
  return lines.join("\n\n") + (lines.length > 0 ? "\n" : "");
}

function subcreator_sanitize_filename_part(value) {
  // // Build readable temp filenames without characters rejected by macOS or Windows file systems.
  var normalized = subcreator_trim_string(String(value || "SubCreator"));
  normalized = normalized.replace(/[\\\/:\*\?"<>\|]+/g, "-").replace(/\s+/g, "-");
  if (!normalized) {
    normalized = "SubCreator";
  }
  return normalized.slice(0, 48);
}

function subcreator_get_project_adjacent_srt_folder() {
  // // Prefer a visible SRT folder next to the current .prproj so generated subtitle sources stay with the project.
  var projectPath = "";
  try {
    projectPath = subcreator_trim_string(app && app.project && app.project.path ? String(app.project.path) : "");
  } catch (projectPathError) {
    projectPath = "";
  }

  if (!projectPath) {
    return null;
  }

  var projectFile = new File(projectPath);
  var projectFolder = projectFile.parent;
  if (!projectFolder) {
    return null;
  }

  var srtFolder = new Folder(projectFolder.fsName + "/SRT");
  if (!srtFolder.exists && !srtFolder.create()) {
    throw new Error("Unable to create project SRT folder: " + srtFolder.fsName);
  }
  return srtFolder;
}

function subcreator_get_native_subtitle_folder() {
  // // Store generated SRT sources beside the Premiere project, with a user-data fallback for unsaved projects.
  var projectSrtFolder = subcreator_get_project_adjacent_srt_folder();
  if (projectSrtFolder) {
    return projectSrtFolder;
  }

  var subcreatorFolder = new Folder(Folder.userData.fsName + "/SubCreator");
  if (!subcreatorFolder.exists) {
    subcreatorFolder.create();
  }
  var baseFolder = new Folder(subcreatorFolder.fsName + "/native-subtitles");
  if (!baseFolder.exists && !baseFolder.create()) {
    throw new Error("Unable to create native subtitle folder: " + baseFolder.fsName);
  }
  return baseFolder;
}

function subcreator_write_native_subtitle_srt(sequence, cues) {
  // // Write planned cues to a uniquely named SRT that Premiere can import as a captions source clip.
  var srtText = subcreator_serialize_cues_to_srt(cues);
  if (!srtText) {
    throw new Error("No valid subtitle cues to import.");
  }

  var targetFolder = subcreator_get_native_subtitle_folder();
  var sequenceName = sequence && sequence.name ? String(sequence.name) : "Sequence";
  var fileName = "SubCreator-" + subcreator_sanitize_filename_part(sequenceName) + "-" + String(new Date().getTime()) + ".srt";
  var fileRef = new File(targetFolder.fsName + "/" + fileName);
  fileRef.encoding = "UTF-8";
  fileRef.lineFeed = "Windows";
  if (!fileRef.open("w")) {
    throw new Error("Unable to write native subtitle SRT: " + fileRef.fsName);
  }
  fileRef.write(srtText);
  fileRef.close();
  return fileRef.fsName;
}

function subcreator_normalize_compare_path(value) {
  // // Normalize host paths before comparing imported project items with the generated SRT file.
  return subcreator_normalize_system_path(String(value || "")).replace(/\\/g, "/").toLowerCase();
}

function subcreator_find_project_item_by_path(rootItem, filePath, fileName) {
  // // Walk the project tree to find the SRT ProjectItem Premiere just imported.
  if (!rootItem) {
    return null;
  }

  var targetPath = subcreator_normalize_compare_path(filePath);
  var targetName = String(fileName || "").toLowerCase();
  var children = subcreator_collection_to_array(rootItem.children);
  for (var i = 0; i < children.length; i += 1) {
    var child = children[i];
    if (!child) {
      continue;
    }

    var childPath = "";
    try {
      childPath = typeof child.getMediaPath === "function" ? child.getMediaPath() : "";
    } catch (pathError) {
      childPath = "";
    }
    if (childPath && subcreator_normalize_compare_path(childPath) === targetPath) {
      return child;
    }

    var childName = "";
    try {
      childName = String(child.name || "").toLowerCase();
    } catch (nameError) {
      childName = "";
    }
    if (targetName && childName === targetName) {
      var namePath = subcreator_normalize_compare_path(childPath);
      if (!namePath || namePath === targetPath) {
        return child;
      }
    }

    var nested = subcreator_find_project_item_by_path(child, filePath, fileName);
    if (nested) {
      return nested;
    }
  }

  return null;
}

function subcreator_is_project_bin_item(projectItem) {
  // // Detect Premiere project bins without depending on one host version's enum availability.
  if (!projectItem) {
    return false;
  }

  try {
    if (typeof ProjectItemType !== "undefined" && typeof ProjectItemType.BIN !== "undefined" && projectItem.type === ProjectItemType.BIN) {
      return true;
    }
  } catch (typeError) {}

  try {
    return Boolean(projectItem.children && typeof projectItem.createBin === "function");
  } catch (fallbackError) {
    return false;
  }
}

function subcreator_find_child_bin_by_name(parentItem, binName) {
  // // Reuse an existing top-level bin so repeated native subtitle runs stay grouped in Premiere.
  if (!parentItem) {
    return null;
  }

  var targetName = String(binName || "").toLowerCase();
  var children = subcreator_collection_to_array(parentItem.children);
  for (var i = 0; i < children.length; i += 1) {
    var child = children[i];
    var childName = "";
    try {
      childName = String(child && child.name ? child.name : "").toLowerCase();
    } catch (nameError) {
      childName = "";
    }

    if (childName === targetName && subcreator_is_project_bin_item(child)) {
      return child;
    }
  }

  return null;
}

function subcreator_get_or_create_project_bin(binName) {
  // // Import generated subtitle source files into a dedicated Premiere project bin instead of the root.
  var rootItem = app && app.project ? app.project.rootItem : null;
  if (!rootItem) {
    return null;
  }

  var existingBin = subcreator_find_child_bin_by_name(rootItem, binName);
  if (existingBin) {
    return existingBin;
  }

  try {
    if (typeof rootItem.createBin === "function") {
      var createdBin = rootItem.createBin(binName);
      if (createdBin && subcreator_is_project_bin_item(createdBin)) {
        return createdBin;
      }
    }
  } catch (createError) {}

  return subcreator_find_child_bin_by_name(rootItem, binName) || rootItem;
}

function subcreator_import_native_subtitle_project_item(srtPath) {
  // // Import the generated SRT into the project and return the ProjectItem needed by createCaptionTrack.
  var fileRef = new File(srtPath);
  if (!fileRef.exists) {
    throw new Error("Native subtitle SRT not found: " + srtPath);
  }

  var targetBin = subcreator_get_or_create_project_bin("SRT") || app.project.rootItem;
  var importResult = app.project.importFiles([fileRef.fsName], true, targetBin, false);
  if (!importResult) {
    throw new Error("Premiere could not import native subtitle SRT: " + fileRef.fsName);
  }

  var projectItem = subcreator_find_project_item_by_path(app.project.rootItem, fileRef.fsName, fileRef.name);
  if (!projectItem) {
    throw new Error("Imported native subtitle SRT project item was not found.");
  }
  return projectItem;
}

function subcreator_create_native_caption_track(sequence, projectItem) {
  // // Prefer Premiere's Subtitle caption format, falling back to the API default when the enum is unavailable.
  if (!sequence || typeof sequence.createCaptionTrack !== "function") {
    throw new Error("This Premiere version does not expose createCaptionTrack().");
  }

  try {
    if (typeof Sequence !== "undefined" && typeof Sequence.CAPTION_FORMAT_SUBTITLE !== "undefined") {
      return sequence.createCaptionTrack(projectItem, 0, Sequence.CAPTION_FORMAT_SUBTITLE);
    }
  } catch (enumError) {}

  return sequence.createCaptionTrack(projectItem, 0);
}

function subcreator_apply_native_subtitles(payloadEncoded) {
  // // Create one native Premiere subtitle track by importing generated cue timing as an SRT source clip.
  try {
    var payloadText = subcreator_decode_payload(payloadEncoded);
    var payload = JSON.parse(payloadText);

    if (!app || !app.project || !app.project.activeSequence) {
      return JSON.stringify({ ok: false, error: "No active sequence in Premiere." });
    }

    var sequence = app.project.activeSequence;
    var sequenceIdentity = subcreator_get_sequence_identity(sequence);
    var cues = payload.cues || [];
    var srtPath = subcreator_write_native_subtitle_srt(sequence, cues);
    var projectItem = subcreator_import_native_subtitle_project_item(srtPath);
    var created = subcreator_create_native_caption_track(sequence, projectItem);

    return JSON.stringify({
      ok: Boolean(created),
      totalCues: cues.length,
      insertedNativeSubtitles: Boolean(created) ? cues.length : 0,
      nativeSubtitleTrackCreated: Boolean(created),
      nativeSubtitleSrtPath: srtPath,
      importedProjectItemName: projectItem && projectItem.name ? String(projectItem.name) : "",
      projectDocumentId: sequenceIdentity.projectDocumentId,
      projectPath: sequenceIdentity.projectPath,
      sequenceID: sequenceIdentity.sequenceID,
      sequenceName: sequenceIdentity.sequenceName
    });
  } catch (error) {
    return JSON.stringify({ ok: false, error: error.toString() });
  }
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
    var sequenceFrameDurationSeconds = subcreator_get_sequence_frame_duration_seconds(sequence);
    var sequenceIdentity = subcreator_get_sequence_identity(sequence);
    var options = payload.options || {};
    var cues = payload.cues || [];
    var templateTextPayloads = subcreator_decode_template_text_payloads(options.premiereTemplateTextPayloads || []);

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
    var bakedTextValidated = 0;
    var bakedTextRetries = 0;
    var bakedTextFailed = 0;
    var debugLines = [];

    for (var i = 0; i < cues.length; i += 1) {
      var cue = cues[i];
      var startSeconds = Number(cue.startSeconds);
      var endSeconds = Number(cue.endSeconds);
      var nextCue = i + 1 < cues.length ? cues[i + 1] || {} : null;
      var requestedStartSeconds = startSeconds;
      var requestedEndSeconds = endSeconds;
      if (hasMogrt) {
        startSeconds = subcreator_snap_seconds_to_nearest_frame(startSeconds, sequenceFrameDurationSeconds);
        endSeconds = subcreator_snap_seconds_to_nearest_frame(endSeconds, sequenceFrameDurationSeconds);
        var nextCueStartSeconds = nextCue
          ? subcreator_snap_seconds_to_nearest_frame(Number(nextCue.startSeconds), sequenceFrameDurationSeconds)
          : NaN;
        endSeconds = subcreator_snap_nearby_caption_end(
          endSeconds,
          nextCueStartSeconds,
          sequenceFrameDurationSeconds
        );
        if (endSeconds <= startSeconds) {
          endSeconds = startSeconds + sequenceFrameDurationSeconds;
        }
      }
      var text = cue.text || "";
      var cueDebugEnabled = i < 4;
      var cueDebugPrefix = cueDebugEnabled
        ? "cue=" + String(i) + " start=" + String(startSeconds) + " end=" + String(endSeconds) + " text=" + subcreator_visual_preview_debug_value(text, 80)
        : "";

      if (cueDebugEnabled && i === 0) {
        subcreator_debug_push_limited(debugLines, "template_text_payloads=" + String(templateTextPayloads.length), 120);
      }
      if (cueDebugEnabled && (startSeconds !== requestedStartSeconds || endSeconds !== requestedEndSeconds)) {
        subcreator_debug_push_limited(
          debugLines,
          cueDebugPrefix +
            " frame_aligned=" +
            String(requestedStartSeconds) +
            "-" +
            String(requestedEndSeconds) +
            " -> " +
            String(startSeconds) +
            "-" +
            String(endSeconds),
          120
        );
      }

      if (hasMogrt && typeof sequence.importMGT === "function") {
        var cueMogrtPath = subcreator_trim_string(String(cue.mogrtPathOverride || ""));
        var cuePathCandidates = cueMogrtPath ? subcreator_build_mogrt_path_candidates(cueMogrtPath) : pathCandidates;
        var cueStyleConfig = subcreator_clone_style_config_with_clip_duration(options.style || {}, endSeconds - startSeconds);
        if (cueDebugEnabled && cueMogrtPath) {
          subcreator_debug_push_limited(debugLines, cueDebugPrefix + " override_path=" + cueMogrtPath, 120);
        }
        var importAttempt = subcreator_try_import_mogrt(
          sequence,
          cuePathCandidates,
          startSeconds,
          videoTrackIndex,
          audioTrackIndex,
          cueDebugEnabled ? debugLines : null,
          cueDebugPrefix
        );
        mogrtAttempted += importAttempt.attempted;
        if (importAttempt.trackItem) {
          var durationApplied = false;
          if (Boolean(cue.skipTextApply)) {
            var bakedValidationDebug = cueDebugEnabled ? debugLines : debugLines.length < 120 ? debugLines : null;
            var bakedValidationPrefix = cueDebugPrefix || "cue=" + String(i);
            var bakedTextValid = subcreator_mogrt_has_nonempty_text_property(
              importAttempt.trackItem,
              bakedValidationDebug,
              bakedValidationPrefix
            );
            // // Premiere can briefly import a pre-baked MOGRT without its text layer during long batches, so retry with progressive backoff before falling back to a marker.
            var bakedRetryDelaysMs = [220, 460, 860, 1400, 2200];
            var bakedRetryIndex = 0;
            while (!bakedTextValid && bakedRetryIndex < bakedRetryDelaysMs.length) {
              bakedTextRetries += 1;
              subcreator_remove_track_item_without_ripple(importAttempt.trackItem);
              subcreator_sleep_ms(bakedRetryDelaysMs[bakedRetryIndex]);
              importAttempt = subcreator_try_import_mogrt(
                sequence,
                cuePathCandidates,
                startSeconds,
                videoTrackIndex,
                audioTrackIndex,
                bakedValidationDebug,
                bakedValidationPrefix + " retry=" + String(bakedRetryIndex + 1)
              );
              mogrtAttempted += importAttempt.attempted;
              if (!importAttempt.trackItem) {
                break;
              }
              bakedTextValid = subcreator_mogrt_has_nonempty_text_property(
                importAttempt.trackItem,
                bakedValidationDebug,
                bakedValidationPrefix + " retry=" + String(bakedRetryIndex + 1)
              );
              bakedRetryIndex += 1;
            }

            if (!bakedTextValid) {
              bakedTextFailed += 1;
              if (importAttempt.trackItem) {
                subcreator_remove_track_item_without_ripple(importAttempt.trackItem);
              }
              importAttempt.trackItem = null;
            } else {
              bakedTextValidated += 1;
            }
          }
        }

        if (importAttempt.trackItem) {
          insertedMogrt += 1;
          lastImportMode = importAttempt.usedTimeMode;
          lastImportPath = importAttempt.usedPath;
          if (durationApplied) {
            durationAdjusted += 1;
          }
          var controlStats = subcreator_try_set_mogrt_controls(
            importAttempt.trackItem,
            text,
            options.style ? options.style.animationMode : "line",
            cueStyleConfig,
            templateTextPayloads,
            Boolean(cue.skipTextApply),
            cueDebugEnabled ? debugLines : null,
            cueDebugPrefix
          );
          if (Boolean(cue.skipTextApply)) {
            updatedText += 1;
          } else if (controlStats.textUpdates > 0) {
            updatedText += controlStats.textUpdates;
          }
          if (controlStats.animationUpdates > 0) {
            updatedAnimation += controlStats.animationUpdates;
          }
          if (controlStats.layoutUpdates > 0) {
            updatedLayout += controlStats.layoutUpdates;
          }
          var durationResult = subcreator_try_razor_mogrt_duration(
            sequence,
            importAttempt.trackItem,
            videoTrackIndex,
            startSeconds,
            endSeconds,
            cueDebugEnabled ? debugLines : null,
            cueDebugPrefix
          );
          if (durationResult.applied) {
            importAttempt.trackItem = durationResult.trackItem;
            durationApplied = true;
            durationAdjusted += 1;
          }
          subcreator_debug_push_limited(
            cueDebugEnabled ? debugLines : null,
            cueDebugPrefix +
              " finalStart=" +
              String(subcreator_to_seconds(importAttempt.trackItem.start || importAttempt.trackItem.inPoint || importAttempt.trackItem.startTime)) +
              " finalEnd=" +
              String(subcreator_to_seconds(importAttempt.trackItem.end || importAttempt.trackItem.outPoint || importAttempt.trackItem.endTime)),
            120
          );
          // // Give Premiere a small breather between pre-baked imports so Essential Graphics has time to materialize each text layer.
          subcreator_sleep_ms(Boolean(cue.skipTextApply) ? 90 : 20);
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
      mogrtBakedTextValidated: bakedTextValidated,
      mogrtBakedTextRetries: bakedTextRetries,
      mogrtBakedTextFailed: bakedTextFailed,
      videoTrackCreated: videoTrackInfo.created,
      videoTracksBefore: videoTrackInfo.beforeTracks,
      videoTracksAfter: videoTrackInfo.afterTracks,
      videoTrackUsed: videoTrackIndex,
      audioTrackUsed: audioTrackIndex,
      projectDocumentId: sequenceIdentity.projectDocumentId,
      projectPath: sequenceIdentity.projectPath,
      sequenceID: sequenceIdentity.sequenceID,
      sequenceName: sequenceIdentity.sequenceName,
      debug: debugLines
    });
  } catch (error) {
    return JSON.stringify({ ok: false, error: error.toString() });
  }
}
