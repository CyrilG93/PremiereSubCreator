#!/bin/bash
set -eu

# // Install the connected macOS payload into the active user's profile even though Installer runs this script as root.
SUBCREATOR_SCRIPT_DIR="${SUBCREATOR_INSTALLER_SCRIPT_DIR:-$(cd "$(dirname "$0")" && pwd)}"
SUBCREATOR_PAYLOAD_ROOT="${SUBCREATOR_PAYLOAD_ROOT:-${SUBCREATOR_SCRIPT_DIR}/payload}"
SUBCREATOR_RUNTIME_ENV="${SUBCREATOR_RUNTIME_ENV:-${SUBCREATOR_SCRIPT_DIR}/runtime.env}"

if [ ! -f "${SUBCREATOR_RUNTIME_ENV}" ]; then
  echo "Runtime metadata is missing: ${SUBCREATOR_RUNTIME_ENV}" >&2
  exit 1
fi

# // Load generated, shell-escaped runtime metadata embedded by the macOS package builder.
. "${SUBCREATOR_RUNTIME_ENV}"

subcreator_resolve_user() {
  # // Resolve the graphical login user instead of writing files into root's home directory.
  SUBCREATOR_USER="${SUBCREATOR_TEST_USER:-$(stat -f "%Su" /dev/console 2>/dev/null || true)}"
  if [ -z "${SUBCREATOR_USER}" ] || [ "${SUBCREATOR_USER}" = "root" ] || [ "${SUBCREATOR_USER}" = "loginwindow" ]; then
    SUBCREATOR_USER="${SUDO_USER:-}"
  fi
  if [ -z "${SUBCREATOR_USER}" ] || [ "${SUBCREATOR_USER}" = "root" ]; then
    echo "Unable to resolve the macOS login user." >&2
    exit 1
  fi

  SUBCREATOR_UID="$(id -u "${SUBCREATOR_USER}")"
  SUBCREATOR_GID="$(id -g "${SUBCREATOR_USER}")"
  SUBCREATOR_HOME="${SUBCREATOR_TEST_HOME:-$(dscl . -read "/Users/${SUBCREATOR_USER}" NFSHomeDirectory 2>/dev/null | awk '{$1=""; sub(/^ /, ""); print}' || true)}"
  if [ -z "${SUBCREATOR_HOME}" ]; then
    SUBCREATOR_HOME="$(eval echo "~${SUBCREATOR_USER}")"
  fi
}

subcreator_run_as_user() {
  # // Run per-user tools in the graphical user's launch context when available.
  if [ "$(id -u)" -ne 0 ] || [ "${SUBCREATOR_USER}" = "$(id -un)" ]; then
    HOME="${SUBCREATOR_HOME}" USER="${SUBCREATOR_USER}" "$@"
    return
  fi

  if command -v launchctl >/dev/null 2>&1; then
    launchctl asuser "${SUBCREATOR_UID}" sudo -H -u "${SUBCREATOR_USER}" "$@"
  else
    sudo -H -u "${SUBCREATOR_USER}" "$@"
  fi
}

subcreator_json_escape() {
  # // Escape runtime paths before writing the JSON file consumed by the CEP panel.
  printf "%s" "$1" | sed 's/\\/\\\\/g; s/"/\\"/g'
}

subcreator_extract_bundled_template_paths() {
  # // Read the installed catalog so only user-added MOGRT files survive an extension refresh.
  catalog_path="$1"
  output_path="$2"
  : >"${output_path}"
  if [ ! -f "${catalog_path}" ]; then
    return 0
  fi

  sed -n 's/.*"relativePath"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' "${catalog_path}" |
    while IFS= read -r relative_path; do
      normalized_path="${relative_path#./}"
      normalized_path="${normalized_path#/}"
      if [ -n "${normalized_path}" ]; then
        printf "%s\n" "${normalized_path}" >>"${output_path}"
      fi
    done
}

subcreator_backup_templates_from_extension() {
  # // Copy custom templates from one installed extension directory into the shared temporary backup.
  extension_dir="$1"
  backup_root="$2"
  source_root="${extension_dir}/templates/mogrt"
  if [ ! -d "${source_root}" ]; then
    return 0
  fi

  catalog_list="${backup_root}/bundled-$(basename "${extension_dir}").txt"
  subcreator_extract_bundled_template_paths "${extension_dir}/assets/mogrt-catalog.json" "${catalog_list}"
  find "${source_root}" -type f -print0 |
    while IFS= read -r -d '' source_path; do
      relative_path="${source_path#${source_root}/}"
      if [ -s "${catalog_list}" ] && grep -Fqx "${relative_path}" "${catalog_list}" >/dev/null 2>&1; then
        continue
      fi
      target_path="${backup_root}/mogrt/${relative_path}"
      mkdir -p "$(dirname "${target_path}")"
      cp -p "${source_path}" "${target_path}"
    done
}

subcreator_install_extension() {
  # // Replace the extension while preserving custom templates from current and legacy CEP identifiers.
  source_dir="${SUBCREATOR_PAYLOAD_ROOT}/dist/com.cyrilplugin.subcreator"
  dest_dir="${SUBCREATOR_HOME}/Library/Application Support/Adobe/CEP/extensions/com.cyrilplugin.subcreator"
  legacy_dest_dir="${SUBCREATOR_HOME}/Library/Application Support/Adobe/CEP/extensions/com.cyrilg93.subcreator"
  if [ ! -d "${source_dir}" ]; then
    echo "Extension payload is missing: ${source_dir}" >&2
    exit 1
  fi

  backup_root="$(mktemp -d "${TMPDIR:-/tmp}/subcreator-mogrt.XXXXXX")"
  subcreator_backup_templates_from_extension "${dest_dir}" "${backup_root}"
  subcreator_backup_templates_from_extension "${legacy_dest_dir}" "${backup_root}"

  mkdir -p "$(dirname "${dest_dir}")"
  replacement_dir="${dest_dir}.new.$$"
  old_dir="${dest_dir}.old.$$"
  ditto "${source_dir}" "${replacement_dir}"
  if [ -e "${dest_dir}" ]; then
    mv "${dest_dir}" "${old_dir}"
  fi
  mv "${replacement_dir}" "${dest_dir}"

  if [ -d "${backup_root}/mogrt" ]; then
    mkdir -p "${dest_dir}/templates/mogrt"
    rsync -a --ignore-existing "${backup_root}/mogrt/" "${dest_dir}/templates/mogrt/"
  fi

  [ ! -e "${old_dir}" ] || rm -rf "${old_dir}"
  [ ! -e "${legacy_dest_dir}" ] || rm -rf "${legacy_dest_dir}"
  rm -rf "${backup_root}"
  chown -R "${SUBCREATOR_UID}:${SUBCREATOR_GID}" "${dest_dir}"
  echo "Sub Creator installed to ${dest_dir}."
}

subcreator_enable_cep_debug_mode() {
  # // Enable unsigned CEP extensions for recent Adobe hosts in the active user's preferences.
  csxs_version=7
  while [ "${csxs_version}" -le 20 ]; do
    subcreator_run_as_user defaults write "com.adobe.CSXS.${csxs_version}" PlayerDebugMode -string "1" >/dev/null 2>&1 || true
    csxs_version=$((csxs_version + 1))
  done
  echo "CEP debug mode enabled for CSXS.7 to CSXS.20."
}

subcreator_install_fonts() {
  # // Install changed bundled fonts while leaving identical user font files untouched.
  fonts_dir="${SUBCREATOR_PAYLOAD_ROOT}/Fonts"
  target_dir="${SUBCREATOR_HOME}/Library/Fonts"
  installed=0
  skipped=0
  failed=0
  if [ ! -d "${fonts_dir}" ]; then
    return 0
  fi

  mkdir -p "${target_dir}"
  while IFS= read -r -d '' font_path; do
    target_path="${target_dir}/$(basename "${font_path}")"
    if [ -f "${target_path}" ]; then
      source_hash="$(shasum -a 256 "${font_path}" | awk '{print tolower($1)}')"
      target_hash="$(shasum -a 256 "${target_path}" | awk '{print tolower($1)}')"
      if [ "${source_hash}" = "${target_hash}" ]; then
        skipped=$((skipped + 1))
        continue
      fi
    fi

    if cp -f "${font_path}" "${target_path}"; then
      chown "${SUBCREATOR_UID}:${SUBCREATOR_GID}" "${target_path}"
      installed=$((installed + 1))
    else
      # // Keep the extension and runtime installation usable if one existing font cannot be replaced.
      failed=$((failed + 1))
      echo "WARNING: unable to update bundled font ${target_path}." >&2
    fi
  done < <(find "${fonts_dir}" -type f \( -iname "*.ttf" -o -iname "*.otf" -o -iname "*.ttc" -o -iname "*.dfont" \) -print0)
  if [ "${installed}" -gt 0 ]; then
    echo "Installed ${installed} bundled font(s) for ${SUBCREATOR_USER}."
  fi
  if [ "${skipped}" -gt 0 ]; then
    echo "Kept ${skipped} identical bundled font(s) already installed."
  fi
  if [ "${failed}" -gt 0 ]; then
    echo "WARNING: kept ${failed} existing font file(s) that could not be updated." >&2
  fi
}

subcreator_runtime_is_current() {
  # // Reuse the installed runtime only after its version, imports, and FFmpeg executable all validate.
  runtime_dir="$1"
  version_file="${runtime_dir}/.subcreator-runtime-version"
  [ -f "${version_file}" ] || return 1
  [ "$(tr -d '\r\n' <"${version_file}")" = "${SUBCREATOR_RUNTIME_VERSION}" ] || return 1
  [ -x "${runtime_dir}/python/bin/python3" ] || return 1
  [ -x "${runtime_dir}/ffmpeg/bin/ffmpeg" ] || return 1
  subcreator_run_as_user "${runtime_dir}/python/bin/python3" -c "import whisper; import whisperx" >/dev/null 2>&1 || return 1
  subcreator_run_as_user "${runtime_dir}/ffmpeg/bin/ffmpeg" -version >/dev/null 2>&1 || return 1
  return 0
}

subcreator_download_runtime() {
  # // Prefer the bundled runtime archive and download it only for explicitly connected-only packages.
  runtime_dir="$1"
  temp_root="$(mktemp -d "${TMPDIR:-/tmp}/subcreator-runtime.XXXXXX")"
  bundled_archive_path="${SUBCREATOR_SCRIPT_DIR}/runtime/${SUBCREATOR_RUNTIME_ASSET_NAME}"
  archive_path="${bundled_archive_path}"
  extracted_root="${temp_root}/extracted"
  mkdir -p "${extracted_root}"

  if [ -f "${bundled_archive_path}" ]; then
    echo "Using the bundled private Whisper runtime for ${SUBCREATOR_RUNTIME_ARCH}..."
  else
    archive_path="${temp_root}/${SUBCREATOR_RUNTIME_ASSET_NAME}"
    echo "Downloading the private Whisper runtime for ${SUBCREATOR_RUNTIME_ARCH}..."
    curl --fail --location --retry 3 --connect-timeout 30 \
      "${SUBCREATOR_RUNTIME_URL}" --output "${archive_path}"
  fi
  actual_hash="$(shasum -a 256 "${archive_path}" | awk '{print tolower($1)}')"
  if [ "${actual_hash}" != "${SUBCREATOR_RUNTIME_SHA256}" ]; then
    rm -rf "${temp_root}"
    echo "Runtime SHA-256 mismatch." >&2
    exit 1
  fi

  tar -xzf "${archive_path}" -C "${extracted_root}"
  new_runtime="${extracted_root}/runtime"
  if [ ! -x "${new_runtime}/python/bin/python3" ] || [ ! -x "${new_runtime}/ffmpeg/bin/ffmpeg" ]; then
    rm -rf "${temp_root}"
    echo "The downloaded runtime archive is incomplete." >&2
    exit 1
  fi

  mkdir -p "$(dirname "${runtime_dir}")"
  old_runtime="${runtime_dir}.old.$$"
  if [ -e "${runtime_dir}" ]; then
    mv "${runtime_dir}" "${old_runtime}"
  fi
  mv "${new_runtime}" "${runtime_dir}"
  [ ! -e "${old_runtime}" ] || rm -rf "${old_runtime}"
  rm -rf "${temp_root}"
  chown -R "${SUBCREATOR_UID}:${SUBCREATOR_GID}" "${runtime_dir}"
  echo "Private runtime installed to ${runtime_dir}."
}

subcreator_validate_runtime() {
  # // Validate Python imports and FFmpeg before exposing the runtime to the extension.
  runtime_dir="$1"
  python_path="${runtime_dir}/python/bin/python3"
  ffmpeg_path="${runtime_dir}/ffmpeg/bin/ffmpeg"
  subcreator_run_as_user "${python_path}" -c "import whisper; import whisperx; print('Whisper runtime validation OK')"
  subcreator_run_as_user "${ffmpeg_path}" -version >/dev/null
}

subcreator_write_runtime_config() {
  # // Persist exact private-runtime paths in the location already read by the panel and host bridge.
  runtime_dir="$1"
  config_dir="${SUBCREATOR_HOME}/Library/Application Support/SubCreator"
  config_file="${config_dir}/subcreator-runtime.json"
  python_path="${runtime_dir}/python/bin/python3"
  ffmpeg_path="${runtime_dir}/ffmpeg/bin/ffmpeg"
  python_version="$("${python_path}" --version 2>&1 | sed -n '1p')"
  generated_at="$(date -u +"%Y-%m-%dT%H:%M:%SZ")"

  mkdir -p "${config_dir}"
  cat >"${config_file}" <<EOF
{
  "version": 1,
  "generatedBy": "subcreator_install_macos_private_runtime.sh",
  "generatedAtUtc": "$(subcreator_json_escape "${generated_at}")",
  "pythonCommand": "$(subcreator_json_escape "${python_path}")",
  "pythonLabel": "Sub Creator private Python",
  "pythonPath": "$(subcreator_json_escape "${python_path}")",
  "pythonVersion": "$(subcreator_json_escape "${python_version}")",
  "whisperPath": "",
  "ffmpegPath": "$(subcreator_json_escape "${ffmpeg_path}")",
  "pathHints": [
    "$(subcreator_json_escape "${runtime_dir}/python/bin")",
    "$(subcreator_json_escape "${runtime_dir}/ffmpeg/bin")",
    "/usr/bin",
    "/bin"
  ]
}
EOF
  chmod 600 "${config_file}"
  chown -R "${SUBCREATOR_UID}:${SUBCREATOR_GID}" "${config_dir}"
  echo "Runtime config written: ${config_file}"
}

subcreator_resolve_user

SUBCREATOR_RUNTIME_DIR="${SUBCREATOR_HOME}/Library/Application Support/SubCreator/runtime"
if [ "${SUBCREATOR_RUNTIME_ARCH}" != "$(uname -m)" ]; then
  echo "Installer runtime architecture ${SUBCREATOR_RUNTIME_ARCH} does not match this Mac ($(uname -m))." >&2
  exit 1
fi

subcreator_install_extension
if [ "${SUBCREATOR_SKIP_CEP_DEBUG:-0}" != "1" ]; then
  subcreator_enable_cep_debug_mode
fi
subcreator_install_fonts

if subcreator_runtime_is_current "${SUBCREATOR_RUNTIME_DIR}"; then
  echo "Keeping the compatible private runtime already installed."
else
  if [ "${SUBCREATOR_SKIP_RUNTIME_DOWNLOAD:-0}" = "1" ]; then
    echo "Runtime download skipped by test configuration."
  else
    subcreator_download_runtime "${SUBCREATOR_RUNTIME_DIR}"
  fi
fi

if [ "${SUBCREATOR_SKIP_RUNTIME_DOWNLOAD:-0}" != "1" ]; then
  subcreator_validate_runtime "${SUBCREATOR_RUNTIME_DIR}"
  printf "%s\n" "${SUBCREATOR_RUNTIME_VERSION}" >"${SUBCREATOR_RUNTIME_DIR}/.subcreator-runtime-version"
  chown "${SUBCREATOR_UID}:${SUBCREATOR_GID}" "${SUBCREATOR_RUNTIME_DIR}/.subcreator-runtime-version"
  subcreator_write_runtime_config "${SUBCREATOR_RUNTIME_DIR}"
fi

echo "Installation complete. Restart Premiere Pro, then open Window > Extensions > Sub Creator."
