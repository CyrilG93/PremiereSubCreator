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

function subcreator_is_default_caption_label(text) {
  // // Detect synthetic/default caption names returned by some Premiere APIs.
  var normalized = subcreator_trim_string(String(text || "")).toLowerCase().replace(/\s+/g, "");
  return normalized === "syntheticcaption";
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

  var syntheticFallback = "";
  var methodNames = ["getCaptionText", "getText", "getSourceText", "getFormattedText", "getTranscriptText"];
  for (var methodIndex = 0; methodIndex < methodNames.length; methodIndex += 1) {
    var methodName = methodNames[methodIndex];
    try {
      if (typeof item[methodName] === "function") {
        var methodText = subcreator_trim_string(String(item[methodName]() || "").replace(/\r/g, "\n"));
        if (methodText) {
          if (!subcreator_is_default_caption_label(methodText)) {
            return methodText;
          }

          if (!syntheticFallback) {
            syntheticFallback = methodText;
          }
        }
      }
    } catch (methodError) {}
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
      if (metadata && metadata.indexOf("SyntheticCaption") === -1) {
        var metadataMatch = metadata.match(/<xmpDM:logComment>([\s\S]*?)<\/xmpDM:logComment>/);
        if (metadataMatch && metadataMatch[1]) {
          return subcreator_trim_string(String(metadataMatch[1]).replace(/\s+/g, " "));
        }
      }
    }
  } catch (metadataError) {}

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

function subcreator_get_or_create_top_video_track_index(sequence) {
  // // Create a new top video track using non-destructive insertion signatures first.
  var currentTracks = sequence && sequence.videoTracks ? Number(sequence.videoTracks.numTracks || 0) : 0;
  var beforeSignatures = [];

  function resolveTopUsableTrackIndex(trackCollection) {
    // // Prefer an empty top track; otherwise use the current top-most track.
    if (!trackCollection || typeof trackCollection.numTracks !== "number") {
      return 0;
    }

    var totalTracks = Number(trackCollection.numTracks || 0);
    if (totalTracks < 1) {
      return 0;
    }

    for (var topIndex = totalTracks - 1; topIndex >= 0; topIndex -= 1) {
      var track = trackCollection[topIndex];
      var clipCount = track && track.clips ? Number(track.clips.numItems || 0) : 0;
      if (clipCount < 1) {
        return topIndex;
      }
    }

    return totalTracks - 1;
  }

  function captureTrackSignatures(trackCollection) {
    var signatures = [];
    if (!trackCollection || typeof trackCollection.numTracks !== "number") {
      return signatures;
    }

    for (var trackIndex = 0; trackIndex < trackCollection.numTracks; trackIndex += 1) {
      var track = trackCollection[trackIndex];
      var clipCount = track && track.clips ? Number(track.clips.numItems || 0) : 0;
      var firstSignature = "";
      var lastSignature = "";

      if (clipCount > 0 && track && track.clips) {
        var firstClip = track.clips[0];
        var lastClip = track.clips[clipCount - 1];
        firstSignature =
          String(firstClip && firstClip.projectItem ? firstClip.projectItem.name : "") +
          "@" +
          String(subcreator_to_seconds(firstClip ? firstClip.start : 0));
        lastSignature =
          String(lastClip && lastClip.projectItem ? lastClip.projectItem.name : "") +
          "@" +
          String(subcreator_to_seconds(lastClip ? lastClip.start : 0));
      }

      signatures.push([String(clipCount), firstSignature, lastSignature].join("|"));
    }

    return signatures;
  }

  function detectInsertedTrackIndex(before, after) {
    if (after.length !== before.length + 1) {
      return -1;
    }

    for (var insertIndex = 0; insertIndex < after.length; insertIndex += 1) {
      var matches = true;

      for (var beforeIndex = 0; beforeIndex < before.length; beforeIndex += 1) {
        var afterIndex = beforeIndex < insertIndex ? beforeIndex : beforeIndex + 1;
        if (before[beforeIndex] !== after[afterIndex]) {
          matches = false;
          break;
        }
      }

      if (matches) {
        return insertIndex;
      }
    }

    return -1;
  }

  beforeSignatures = captureTrackSignatures(sequence.videoTracks);
  var created = false;
  var createdIndex = -1;

  try {
    if (typeof app.enableQE === "function") {
      app.enableQE();
      if (typeof qe !== "undefined" && qe.project && typeof qe.project.getActiveSequence === "function") {
        var qeSequence = qe.project.getActiveSequence();
        if (qeSequence && typeof qeSequence.addTracks === "function") {
          var inserted = false;

          if (!inserted && currentTracks > 0) {
            try {
              // // Most reliable append signature from QE docs/community examples.
              qeSequence.addTracks(1, currentTracks - 1, 0, 0, 0);
              inserted = true;
            } catch (signatureErrorAppendFull) {}
          }

          if (!inserted && currentTracks > 0) {
            try {
              qeSequence.addTracks(1, currentTracks - 1, 0, 0);
              inserted = true;
            } catch (signatureErrorAppendShort) {}
          }

          if (!inserted && currentTracks > 0) {
            try {
              qeSequence.addTracks(1, currentTracks - 1, 0);
              inserted = true;
            } catch (signatureErrorAppendMinimal) {}
          }

          if (!inserted && currentTracks > 0) {
            try {
              qeSequence.addTracks(1, currentTracks - 1);
              inserted = true;
            } catch (signatureErrorAppendTwoArgs) {}
          }

          try {
            if (!inserted && currentTracks < 1) {
              qeSequence.addTracks(1);
              inserted = true;
            }
          } catch (signatureErrorEmptySequence) {}
        }
      }
    }
  } catch (error) {}

  var updatedTracks = sequence && sequence.videoTracks ? Number(sequence.videoTracks.numTracks || 0) : 0;
  var afterSignatures = captureTrackSignatures(sequence.videoTracks);

  if (updatedTracks > currentTracks) {
    created = true;
    createdIndex = detectInsertedTrackIndex(beforeSignatures, afterSignatures);
    if (createdIndex < 0 || createdIndex > updatedTracks - 1) {
      createdIndex = updatedTracks - 1;
    }
  }

  var fallbackIndex = resolveTopUsableTrackIndex(sequence.videoTracks);

  return {
    index: created ? createdIndex : fallbackIndex,
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
