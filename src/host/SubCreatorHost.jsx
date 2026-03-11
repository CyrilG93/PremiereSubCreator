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
