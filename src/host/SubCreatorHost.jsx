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

    var mogrtPath = options.mogrtPath || "";
    var hasMogrt = false;
    if (mogrtPath && mogrtPath.length > 0) {
      var mogrtFile = new File(mogrtPath);
      hasMogrt = mogrtFile.exists;
    }

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
      mogrtUsed: hasMogrt
    });
  } catch (error) {
    return JSON.stringify({ ok: false, error: error.toString() });
  }
}
