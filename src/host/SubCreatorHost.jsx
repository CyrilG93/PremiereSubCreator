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

function subcreator_resolve_mogrt_path(options) {
  // // Prioritize bundled template path and fallback to manual absolute path.
  var templateRelativePath = options.mogrtTemplateRelativePath || "";
  if (templateRelativePath && templateRelativePath.length > 0) {
    var extensionRoot = subcreator_resolve_extension_root();
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

  return "";
}

function subcreator_get_file_extension(input) {
  // // Extract file extension in lowercase without leading dot.
  var value = String(input || "");
  var dotIndex = value.lastIndexOf(".");
  if (dotIndex < 0 || dotIndex === value.length - 1) {
    return "";
  }

  return value.substring(dotIndex + 1).toLowerCase();
}

function subcreator_is_caption_extension(extension) {
  // // Keep only subtitle-related file extensions from project bins.
  var ext = String(extension || "").toLowerCase();
  return (
    ext === "srt" ||
    ext === "vtt" ||
    ext === "stl" ||
    ext === "scc" ||
    ext === "mcc" ||
    ext === "itt" ||
    ext === "dfxp"
  );
}

function subcreator_collect_caption_sources(projectItem, currentBinPath, collector) {
  // // Recursively scan project bins and collect caption file assets.
  if (!projectItem) {
    return;
  }

  var mediaPath = "";
  if (typeof projectItem.getMediaPath === "function") {
    try {
      mediaPath = projectItem.getMediaPath();
    } catch (error) {
      mediaPath = "";
    }
  }

  var itemName = String(projectItem.name || "");
  var extension = subcreator_get_file_extension(mediaPath || itemName);

  if (mediaPath && mediaPath !== "0" && subcreator_is_caption_extension(extension)) {
    collector.push({
      id: String(projectItem.nodeId || "") || String(collector.length + 1),
      name: itemName || new File(mediaPath).name,
      mediaPath: String(mediaPath),
      extension: extension,
      binPath: currentBinPath
    });
  }

  var children = projectItem.children;
  if (!children || typeof children.numItems === "undefined") {
    return;
  }

  var nextBinPath = currentBinPath;
  if (itemName && itemName !== "Root") {
    nextBinPath = currentBinPath ? currentBinPath + "/" + itemName : itemName;
  }

  for (var i = 0; i < children.numItems; i += 1) {
    subcreator_collect_caption_sources(children[i], nextBinPath, collector);
  }
}

function subcreator_sort_caption_sources(items) {
  // // Keep source list stable and easy to scan in UI dropdown.
  items.sort(function (left, right) {
    var leftName = String(left.name || "").toLowerCase();
    var rightName = String(right.name || "").toLowerCase();
    if (leftName < rightName) {
      return -1;
    }
    if (leftName > rightName) {
      return 1;
    }
    var leftPath = String(left.mediaPath || "").toLowerCase();
    var rightPath = String(right.mediaPath || "").toLowerCase();
    if (leftPath < rightPath) {
      return -1;
    }
    if (leftPath > rightPath) {
      return 1;
    }
    return 0;
  });
}

function subcreator_list_caption_sources() {
  // // Return caption files currently available in Premiere project bins.
  try {
    if (!app || !app.project || !app.project.rootItem) {
      return subcreator_error("No open project in Premiere.");
    }

    var sources = [];
    subcreator_collect_caption_sources(app.project.rootItem, "", sources);
    subcreator_sort_caption_sources(sources);
    return subcreator_ok(sources);
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_read_text_file(encodedPath) {
  // // Read a text file from disk and return UTF-8-ish content to the panel.
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
    var hasMogrt = mogrtPath && mogrtPath.length > 0;

    var videoTrackIndex = Number(options.videoTrackIndex);
    if (isNaN(videoTrackIndex)) {
      videoTrackIndex = 0;
    }

    var audioTrackIndex = Number(options.audioTrackIndex);
    if (isNaN(audioTrackIndex)) {
      audioTrackIndex = 0;
    }

    var insertedMogrt = 0;
    var insertedMarkers = 0;
    var updatedText = 0;

    for (var i = 0; i < cues.length; i += 1) {
      var cue = cues[i];
      var startSeconds = Number(cue.startSeconds);
      var endSeconds = Number(cue.endSeconds);
      var text = cue.text || "";

      if (hasMogrt && typeof sequence.importMGT === "function") {
        var insertedItem = sequence.importMGT(mogrtPath, startSeconds, videoTrackIndex, audioTrackIndex);
        if (insertedItem) {
          insertedMogrt += 1;
          if (subcreator_try_set_mogrt_text(insertedItem, text)) {
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
      mogrtPathResolved: mogrtPath
    });
  } catch (error) {
    return JSON.stringify({ ok: false, error: error.toString() });
  }
}
