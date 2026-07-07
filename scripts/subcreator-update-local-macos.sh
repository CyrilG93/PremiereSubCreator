#!/bin/bash
set -euo pipefail

# // Copy a freshly built Sub Creator extension into the current user's macOS CEP folder.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
DIST_DIR="${REPO_ROOT}/dist/com.cyrilplugin.subcreator"
DESTINATION="${SUBCREATOR_LOCAL_MACOS_DESTINATION:-${HOME}/Library/Application Support/Adobe/CEP/extensions/com.cyrilplugin.subcreator}"
LEGACY_DESTINATION="${HOME}/Library/Application Support/Adobe/CEP/extensions/com.cyrilg93.subcreator"
SKIP_BUILD=0
DRY_RUN=0
SKIP_CEP_DEBUG=0
SKIP_FONTS=0

subcreator_info() {
  # // Keep local-update output readable when launched from npm or the root helper script.
  printf '[Sub Creator] %s\n' "$1"
}

subcreator_usage() {
  # // Document the small set of options used for local testing workflows.
  cat <<'USAGE'
Usage: scripts/subcreator-update-local-macos.sh [options]

Options:
  --destination <path>  Copy the extension to a custom CEP extension folder.
  --skip-build         Copy the current dist folder without running the build.
  --dry-run            Print the planned actions without changing files.
  --skip-cep-debug     Do not write Adobe CSXS PlayerDebugMode preferences.
  --skip-fonts         Do not refresh bundled fonts in ~/Library/Fonts.
  -h, --help           Show this help.
USAGE
}

while [ "$#" -gt 0 ]; do
  case "$1" in
    --destination)
      if [ "$#" -lt 2 ]; then
        echo "--destination requires a path." >&2
        exit 1
      fi
      DESTINATION="$2"
      shift 2
      ;;
    --skip-build)
      SKIP_BUILD=1
      shift
      ;;
    --dry-run)
      DRY_RUN=1
      shift
      ;;
    --skip-cep-debug)
      SKIP_CEP_DEBUG=1
      shift
      ;;
    --skip-fonts)
      SKIP_FONTS=1
      shift
      ;;
    -h|--help)
      subcreator_usage
      exit 0
      ;;
    *)
      echo "Unknown option: $1" >&2
      subcreator_usage >&2
      exit 1
      ;;
  esac
done

subcreator_build() {
  # // Rebuild the TypeScript panel before copying so local tests include current source edits.
  if [ "${SKIP_BUILD}" -eq 1 ]; then
    subcreator_info "Skipping build because --skip-build was provided."
    return
  fi

  if [ "${DRY_RUN}" -eq 1 ]; then
    subcreator_info "Would run npm run subcreator:build"
    return
  fi

  (cd "${REPO_ROOT}" && npm run subcreator:build)
}

subcreator_extract_bundled_template_paths() {
  # // Read the installed catalog so bundled MOGRT files can be replaced while user-added files are retained.
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
        printf '%s\n' "${normalized_path}" >>"${output_path}"
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
      if [ -s "${catalog_list}" ] && grep -Fqx -- "${relative_path}" "${catalog_list}" >/dev/null 2>&1; then
        continue
      fi
      target_path="${backup_root}/mogrt/${relative_path}"
      mkdir -p "$(dirname "${target_path}")"
      cp -p "${source_path}" "${target_path}"
    done
}

subcreator_copy_local_build() {
  # // Replace the installed CEP extension with the freshly built dist folder while keeping user templates.
  if [ ! -d "${DIST_DIR}" ]; then
    echo "Build missing: ${DIST_DIR}" >&2
    exit 1
  fi

  if [ "${DRY_RUN}" -eq 1 ]; then
    subcreator_info "Would copy ${DIST_DIR} -> ${DESTINATION}"
    subcreator_info "Would remove legacy folder ${LEGACY_DESTINATION} if present"
    return
  fi

  backup_root="$(mktemp -d "${TMPDIR:-/tmp}/subcreator-local-update.XXXXXX")"
  replacement_dir="${DESTINATION}.new.$$"
  old_dir="${DESTINATION}.old.$$"
  trap 'rm -R "${backup_root}" "${replacement_dir}" 2>/dev/null || true' EXIT

  subcreator_backup_templates_from_extension "${DESTINATION}" "${backup_root}"
  subcreator_backup_templates_from_extension "${LEGACY_DESTINATION}" "${backup_root}"

  mkdir -p "$(dirname "${DESTINATION}")"
  ditto "${DIST_DIR}" "${replacement_dir}"
  if [ -e "${DESTINATION}" ]; then
    mv "${DESTINATION}" "${old_dir}"
  fi
  mv "${replacement_dir}" "${DESTINATION}"

  if [ -d "${backup_root}/mogrt" ]; then
    mkdir -p "${DESTINATION}/templates/mogrt"
    rsync -a --ignore-existing "${backup_root}/mogrt/" "${DESTINATION}/templates/mogrt/"
  fi

  if [ -e "${old_dir}" ]; then
    rm -R "${old_dir}"
  fi
  if [ -e "${LEGACY_DESTINATION}" ]; then
    rm -R "${LEGACY_DESTINATION}"
  fi
  rm -R "${backup_root}"
  trap - EXIT
}

subcreator_enable_cep_debug_mode() {
  # // Enable unsigned CEP extensions for current-user Adobe hosts across recent CSXS versions.
  if [ "${SKIP_CEP_DEBUG}" -eq 1 ]; then
    subcreator_info "Skipping CEP debug mode because --skip-cep-debug was provided."
    return
  fi

  if [ "${DRY_RUN}" -eq 1 ]; then
    subcreator_info "Would enable CEP debug mode for CSXS.7 to CSXS.20"
    return
  fi

  csxs_version=7
  while [ "${csxs_version}" -le 20 ]; do
    defaults write "com.adobe.CSXS.${csxs_version}" PlayerDebugMode -string "1" >/dev/null 2>&1 || true
    csxs_version=$((csxs_version + 1))
  done
  subcreator_info "CEP debug mode enabled for CSXS.7 to CSXS.20."
}

subcreator_install_fonts() {
  # // Refresh changed bundled fonts while leaving identical user font files untouched.
  if [ "${SKIP_FONTS}" -eq 1 ]; then
    subcreator_info "Skipping bundled fonts because --skip-fonts was provided."
    return
  fi

  fonts_dir="${REPO_ROOT}/Fonts"
  target_dir="${HOME}/Library/Fonts"
  installed=0
  skipped=0
  failed=0
  if [ ! -d "${fonts_dir}" ]; then
    return
  fi

  if [ "${DRY_RUN}" -eq 1 ]; then
    subcreator_info "Would refresh bundled fonts in ${target_dir}"
    return
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
      installed=$((installed + 1))
    else
      failed=$((failed + 1))
      echo "WARNING: unable to update bundled font ${target_path}." >&2
    fi
  done < <(find "${fonts_dir}" -type f \( -iname "*.ttf" -o -iname "*.otf" -o -iname "*.ttc" -o -iname "*.dfont" \) -print0)

  if [ "${installed}" -gt 0 ]; then
    subcreator_info "Updated ${installed} bundled font(s)."
  fi
  if [ "${skipped}" -gt 0 ]; then
    subcreator_info "Kept ${skipped} identical bundled font(s) already installed."
  fi
  if [ "${failed}" -gt 0 ]; then
    echo "WARNING: kept ${failed} existing font file(s) that could not be updated." >&2
  fi
}

subcreator_info "Updating local CEP plugin from ${REPO_ROOT}"
subcreator_info "Destination: ${DESTINATION}"
subcreator_build
subcreator_copy_local_build
subcreator_enable_cep_debug_mode
subcreator_install_fonts
subcreator_info "Local update complete. Restart Premiere Pro, then open Window > Extensions > Sub Creator."
subcreator_info "Existing private runtime and Whisper models were preserved."
