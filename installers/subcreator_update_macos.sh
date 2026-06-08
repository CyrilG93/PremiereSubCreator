#!/bin/bash
set -eu

# // Install only the architecture-independent Sub Creator update payload for the active macOS user.
SUBCREATOR_SCRIPT_DIR="${SUBCREATOR_INSTALLER_SCRIPT_DIR:-$(cd "$(dirname "$0")" && pwd)}"
SUBCREATOR_PAYLOAD_ROOT="${SUBCREATOR_PAYLOAD_ROOT:-${SUBCREATOR_SCRIPT_DIR}/payload}"

subcreator_resolve_user() {
  # // Resolve the graphical login user because macOS Installer executes package scripts as root.
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
  # // Run preferences commands in the graphical user's launch context when Installer is running as root.
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

subcreator_extract_bundled_template_paths() {
  # // Read the installed catalog so bundled MOGRT files are replaced while user-added files are retained.
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
  # // Copy custom MOGRT files from one installed extension into the temporary update backup.
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
  # // Replace current and legacy CEP installations while restoring custom MOGRT files afterward.
  source_dir="${SUBCREATOR_PAYLOAD_ROOT}/dist/com.cyrilplugin.subcreator"
  dest_dir="${SUBCREATOR_HOME}/Library/Application Support/Adobe/CEP/extensions/com.cyrilplugin.subcreator"
  legacy_dest_dir="${SUBCREATOR_HOME}/Library/Application Support/Adobe/CEP/extensions/com.cyrilg93.subcreator"
  if [ ! -d "${source_dir}" ]; then
    echo "Extension update payload is missing: ${source_dir}" >&2
    exit 1
  fi

  backup_root="$(mktemp -d "${TMPDIR:-/tmp}/subcreator-update-mogrt.XXXXXX")"
  replacement_dir="${dest_dir}.new.$$"
  old_dir="${dest_dir}.old.$$"
  trap 'rm -rf "${backup_root}" "${replacement_dir}"' EXIT

  subcreator_backup_templates_from_extension "${dest_dir}" "${backup_root}"
  subcreator_backup_templates_from_extension "${legacy_dest_dir}" "${backup_root}"

  mkdir -p "$(dirname "${dest_dir}")"
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
  trap - EXIT
  chown -R "${SUBCREATOR_UID}:${SUBCREATOR_GID}" "${dest_dir}"
  echo "Sub Creator update installed to ${dest_dir}."
}

subcreator_enable_cep_debug_mode() {
  # // Keep unsigned CEP extensions enabled for supported Adobe host preference versions.
  csxs_version=7
  while [ "${csxs_version}" -le 20 ]; do
    subcreator_run_as_user defaults write "com.adobe.CSXS.${csxs_version}" PlayerDebugMode -string "1" >/dev/null 2>&1 || true
    csxs_version=$((csxs_version + 1))
  done
  echo "CEP debug mode enabled for CSXS.7 to CSXS.20."
}

subcreator_install_fonts() {
  # // Refresh bundled fonts in the user's Library without touching system font directories.
  fonts_dir="${SUBCREATOR_PAYLOAD_ROOT}/Fonts"
  target_dir="${SUBCREATOR_HOME}/Library/Fonts"
  if [ ! -d "${fonts_dir}" ]; then
    return 0
  fi

  mkdir -p "${target_dir}"
  find "${fonts_dir}" -type f \( -iname "*.ttf" -o -iname "*.otf" -o -iname "*.ttc" -o -iname "*.dfont" \) -print0 |
    while IFS= read -r -d '' font_path; do
      cp -f "${font_path}" "${target_dir}/$(basename "${font_path}")"
    done
  installed="$(find "${fonts_dir}" -type f \( -iname "*.ttf" -o -iname "*.otf" -o -iname "*.ttc" -o -iname "*.dfont" \) | wc -l | tr -d ' ')"
  chown -R "${SUBCREATOR_UID}:${SUBCREATOR_GID}" "${target_dir}"
  if [ "${installed}" -gt 0 ]; then
    echo "Updated ${installed} bundled font(s) for ${SUBCREATOR_USER}."
  fi
}

subcreator_resolve_user
subcreator_install_extension
if [ "${SUBCREATOR_SKIP_CEP_DEBUG:-0}" != "1" ]; then
  subcreator_enable_cep_debug_mode
fi
subcreator_install_fonts

# // Leave the private runtime, its configuration, and every downloaded Whisper model unchanged.
echo "Update complete. Existing runtime and Whisper models were preserved."
echo "Restart Premiere Pro, then open Window > Extensions > Sub Creator."
