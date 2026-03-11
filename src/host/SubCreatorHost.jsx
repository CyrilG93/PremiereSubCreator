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

function subcreator_build_whisper_command(audioPath, outputDir, model, languageCode) {
  // // Build CLI command string for local Whisper execution.
  var whisperBinary = "whisper";
  var modelArg = model && model.length > 0 ? model : "base";
  var languageArg = languageCode && languageCode.length > 0 ? languageCode : "";

  if (subcreator_is_windows()) {
    var cmd =
      whisperBinary +
      " " +
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

  var shellCmd =
    whisperBinary +
    " " +
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

function subcreator_extract_text_from_item(item) {
  // // Read caption text from known item methods/properties.
  if (!item) {
    return "";
  }

  try {
    if (typeof item.getCaptionText === "function") {
      return String(item.getCaptionText() || "");
    }
  } catch (error1) {}

  try {
    if (typeof item.getText === "function") {
      return String(item.getText() || "");
    }
  } catch (error2) {}

  if (typeof item.text !== "undefined") {
    return String(item.text || "");
  }

  if (item.projectItem && item.projectItem.name) {
    return String(item.projectItem.name || "");
  }

  if (item.name) {
    return String(item.name || "");
  }

  return "";
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

        var trackItems =
          subcreator_collection_to_array(selectedTrack.clips) ||
          subcreator_collection_to_array(selectedTrack.items) ||
          subcreator_collection_to_array(selectedTrack.captions);

        var cues = subcreator_extract_cues_from_items(trackItems);
        if (cues.length > 0) {
          return subcreator_ok(cues);
        }
      }
    }

    // // Fallback: try selected timeline items when captionTracks API is unavailable.
    if (typeof sequence.getSelection === "function") {
      var selection = subcreator_collection_to_array(sequence.getSelection());
      var selectionCues = subcreator_extract_cues_from_items(selection);
      if (selectionCues.length > 0) {
        return subcreator_ok(selectionCues);
      }
    }

    return subcreator_error(
      "Impossible de lire la piste caption active avec cette API CEP. Si possible, selectionne les clips caption sur la timeline ou utilise la source SRT."
    );
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_try_set_mogrt_text(trackItem, textValue) {
  // // Attempt to update known text properties on inserted MOGRT components.
  if (!trackItem || !textValue || typeof trackItem.getMGTComponent !== "function") {
    return false;
  }

  var component = trackItem.getMGTComponent();
  if (!component || !component.properties || component.properties.numItems < 1) {
    return false;
  }

  for (var i = 0; i < component.properties.numItems; i += 1) {
    var property = component.properties[i];
    if (!property || !property.displayName || typeof property.setValue !== "function") {
      continue;
    }

    var key = String(property.displayName).toLowerCase();
    if (key.indexOf("source text") !== -1 || key.indexOf("texte source") !== -1 || key === "text") {
      property.setValue(textValue, true);
      return true;
    }
  }

  return false;
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
  // // Prioritize bundled template path and fallback to manual absolute path.
  var extensionRoot = subcreator_resolve_extension_root();
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

  var manualPath = options.mogrtPath || "";
  if (manualPath && manualPath.length > 0) {
    var manualFile = new File(manualPath);
    if (manualFile.exists) {
      return manualFile.fsName;
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

function subcreator_get_or_create_top_video_track_index(sequence) {
  // // Create a new top video track and return its index, with safe fallbacks.
  var currentTracks = sequence && sequence.videoTracks ? Number(sequence.videoTracks.numTracks || 0) : 0;

  try {
    if (typeof app.enableQE === "function") {
      app.enableQE();
      if (typeof qe !== "undefined" && qe.project && typeof qe.project.getActiveSequence === "function") {
        var qeSequence = qe.project.getActiveSequence();
        if (qeSequence && typeof qeSequence.addTracks === "function") {
          if (currentTracks > 0) {
            try {
              qeSequence.addTracks(1, currentTracks - 1, 0);
            } catch (signatureError) {
              qeSequence.addTracks(1);
            }
          } else {
            qeSequence.addTracks(1);
          }
        }
      }
    }
  } catch (error) {}

  var updatedTracks = sequence && sequence.videoTracks ? Number(sequence.videoTracks.numTracks || 0) : 0;
  if (updatedTracks > 0) {
    return updatedTracks - 1;
  }

  return 0;
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

    var videoTrackIndex = subcreator_get_or_create_top_video_track_index(sequence);
    var audioTrackIndex = 0;

    var insertedMogrt = 0;
    var insertedMarkers = 0;
    var updatedText = 0;
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
          if (subcreator_try_set_mogrt_text(importAttempt.trackItem, text)) {
            updatedText += 1;
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
      mogrtUsed: hasMogrt,
      mogrtPathResolved: mogrtPath,
      mogrtPathCandidates: pathCandidates,
      mogrtImportAttempts: mogrtAttempted,
      mogrtLastImportMode: lastImportMode,
      mogrtLastImportPath: lastImportPath,
      videoTrackUsed: videoTrackIndex,
      audioTrackUsed: audioTrackIndex
    });
  } catch (error) {
    return JSON.stringify({ ok: false, error: error.toString() });
  }
}
