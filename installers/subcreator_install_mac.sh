#!/usr/bin/env bash
set -euo pipefail

# // Resolve script and project directories reliably whether the installer is launched from the release root or `installers/`.
SUBCREATOR_SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
if [ -d "${SUBCREATOR_SCRIPT_DIR}/dist/com.cyrilplugin.subcreator" ]; then
  SUBCREATOR_PROJECT_DIR="${SUBCREATOR_SCRIPT_DIR}"
else
  SUBCREATOR_PROJECT_DIR="$(cd "${SUBCREATOR_SCRIPT_DIR}/.." && pwd)"
fi
SUBCREATOR_SOURCE_DIR="${SUBCREATOR_PROJECT_DIR}/dist/com.cyrilplugin.subcreator"
SUBCREATOR_DEST_DIR="${HOME}/Library/Application Support/Adobe/CEP/extensions/com.cyrilplugin.subcreator"
SUBCREATOR_LEGACY_DEST_DIR="${HOME}/Library/Application Support/Adobe/CEP/extensions/com.cyrilg93.subcreator"
SUBCREATOR_RUNTIME_DIR="${HOME}/Library/Application Support/SubCreator"
SUBCREATOR_RUNTIME_FILE="${SUBCREATOR_RUNTIME_DIR}/subcreator-runtime.json"
SUBCREATOR_BUNDLED_MODELS_DIR="${SUBCREATOR_PROJECT_DIR}/Models"
SUBCREATOR_WHISPER_MODELS_CACHE_DIR="${HOME}/.cache/whisper"
SUBCREATOR_PYTHON_CMD=""
SUBCREATOR_PYTHON_VERSION=""
SUBCREATOR_PYTHON_PATH=""
SUBCREATOR_PYTHON_SEEN=""
SUBCREATOR_WHISPER_PATH=""
SUBCREATOR_FFMPEG_PATH=""
SUBCREATOR_PATH_HINTS=""
SUBCREATOR_TEMPLATES_BACKUP_ROOT=""
SUBCREATOR_TEMPLATES_BACKUP_DIR=""

subcreator_enable_cep_debug_mode() {
  # // Enable CEP debug mode for multiple CSXS versions to maximize Adobe host compatibility.
  local csxs_versions=(7 8 9 10 11 12)
  local csxs_version=""
  for csxs_version in "${csxs_versions[@]}"; do
    defaults write "com.adobe.CSXS.${csxs_version}" PlayerDebugMode -string "1" >/dev/null 2>&1 || true
  done
  echo "CEP debug mode enabled for CSXS.7 to CSXS.12"
}

subcreator_append_path_to_profile() {
  # // Append Whisper user-bin directory to a shell profile only once.
  local profile_path="$1"
  local bin_path="$2"
  local export_line="export PATH=\"${bin_path}:\$PATH\""

  if [ ! -f "${profile_path}" ]; then
    touch "${profile_path}"
  fi

  if grep -F "${bin_path}" "${profile_path}" >/dev/null 2>&1; then
    return 1
  fi

  printf "\n# // Added by Sub Creator installer for Whisper CLI\n%s\n" "${export_line}" >>"${profile_path}"
  return 0
}

subcreator_extract_bundled_template_paths() {
  # // Extract previously bundled template relative paths so installer updates only preserve user-added templates.
  local catalog_path="$1"
  local output_path="$2"

  : >"${output_path}"

  if [ ! -f "${catalog_path}" ]; then
    return 0
  fi

  sed -n 's/.*"relativePath"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' "${catalog_path}" |
    while IFS= read -r relative_path; do
      local normalized_path="${relative_path#./}"
      normalized_path="${normalized_path#/}"
      if [ -n "${normalized_path}" ]; then
        printf "%s\n" "${normalized_path}" >>"${output_path}"
      fi
    done
}

subcreator_backup_templates_from_extension() {
  # // Preserve only user-added MOGRT files from one installed extension payload.
  local extension_dir="$1"
  if [ ! -d "${extension_dir}/templates/mogrt" ]; then
    return 0
  fi

  local source_root="${extension_dir}/templates/mogrt"
  local old_catalog_path="${extension_dir}/assets/mogrt-catalog.json"
  if [ -z "${SUBCREATOR_TEMPLATES_BACKUP_ROOT}" ]; then
    SUBCREATOR_TEMPLATES_BACKUP_ROOT="$(mktemp -d "${TMPDIR:-/tmp}/subcreator-mogrt.XXXXXX")"
    SUBCREATOR_TEMPLATES_BACKUP_DIR="${SUBCREATOR_TEMPLATES_BACKUP_ROOT}/mogrt"
    mkdir -p "${SUBCREATOR_TEMPLATES_BACKUP_DIR}"
  fi
  local bundled_paths_file="${SUBCREATOR_TEMPLATES_BACKUP_ROOT}/bundled-relative-paths-$(basename "${extension_dir}").txt"
  subcreator_extract_bundled_template_paths "${old_catalog_path}" "${bundled_paths_file}"

  while IFS= read -r -d '' source_path; do
    local relative_path=""
    local target_path=""
    relative_path="${source_path#${source_root}/}"

    if [ -s "${bundled_paths_file}" ] && grep -Fqx "${relative_path}" "${bundled_paths_file}" >/dev/null 2>&1; then
      continue
    fi

    target_path="${SUBCREATOR_TEMPLATES_BACKUP_DIR}/${relative_path}"
    mkdir -p "$(dirname "${target_path}")"
    cp "${source_path}" "${target_path}"
  done < <(find "${source_root}" -type f -print0)
}

subcreator_backup_existing_templates() {
  # // Preserve custom templates from both the current and legacy CEP folder names before replacement.
  subcreator_backup_templates_from_extension "${SUBCREATOR_DEST_DIR}"
  subcreator_backup_templates_from_extension "${SUBCREATOR_LEGACY_DEST_DIR}"
}

subcreator_restore_existing_templates() {
  # // Merge preserved user-added MOGRT files back without overwriting the freshly installed bundle files.
  if [ -z "${SUBCREATOR_TEMPLATES_BACKUP_DIR}" ] || [ ! -d "${SUBCREATOR_TEMPLATES_BACKUP_DIR}" ]; then
    return 0
  fi

  mkdir -p "${SUBCREATOR_DEST_DIR}/templates/mogrt"
  # // `cp -Rn` can return non-zero on macOS when files are skipped because they already exist, which should not abort install.
  if command -v rsync >/dev/null 2>&1; then
    rsync -a --ignore-existing "${SUBCREATOR_TEMPLATES_BACKUP_DIR}/" "${SUBCREATOR_DEST_DIR}/templates/mogrt/"
  else
    cp -Rn "${SUBCREATOR_TEMPLATES_BACKUP_DIR}/." "${SUBCREATOR_DEST_DIR}/templates/mogrt/" || true
  fi
  if [ -n "${SUBCREATOR_TEMPLATES_BACKUP_ROOT}" ] && [ -d "${SUBCREATOR_TEMPLATES_BACKUP_ROOT}" ]; then
    rm -rf "${SUBCREATOR_TEMPLATES_BACKUP_ROOT}"
  fi
  SUBCREATOR_TEMPLATES_BACKUP_ROOT=""
  SUBCREATOR_TEMPLATES_BACKUP_DIR=""
}

subcreator_copy_bundled_whisper_models() {
  # // Copy bundled Whisper model files into the local cache so the panel can expose a guaranteed starter model without download.
  if [ ! -d "${SUBCREATOR_BUNDLED_MODELS_DIR}" ]; then
    return 0
  fi

  local copied_count=0
  local model_path=""
  local model_paths=("${SUBCREATOR_BUNDLED_MODELS_DIR}"/*.pt)
  local part_path=""
  local part_paths=("${SUBCREATOR_BUNDLED_MODELS_DIR}"/*.pt.part-*)
  local processed_models=""
  if [ "${#model_paths[@]}" -lt 1 ] || [ ! -e "${model_paths[0]}" ]; then
    model_paths=()
  fi

  mkdir -p "${SUBCREATOR_WHISPER_MODELS_CACHE_DIR}"

  for model_path in "${model_paths[@]}"; do
    local model_name=""
    if [ ! -f "${model_path}" ]; then
      continue
    fi

    model_name="$(basename "${model_path}")"
    if [ -f "${SUBCREATOR_WHISPER_MODELS_CACHE_DIR}/${model_name}" ]; then
      continue
    fi

    cp "${model_path}" "${SUBCREATOR_WHISPER_MODELS_CACHE_DIR}/${model_name}"
    copied_count=$((copied_count + 1))
  done

  for part_path in "${part_paths[@]}"; do
    local part_name=""
    local model_name=""
    local ordered_parts=()
    if [ ! -f "${part_path}" ]; then
      continue
    fi

    part_name="$(basename "${part_path}")"
    model_name="${part_name%.part-*}"
    case "|${processed_models}|" in
      *"|${model_name}|"*)
        continue
        ;;
    esac
    processed_models="${processed_models}|${model_name}"

    if [ -f "${SUBCREATOR_WHISPER_MODELS_CACHE_DIR}/${model_name}" ]; then
      continue
    fi

    ordered_parts=("${SUBCREATOR_BUNDLED_MODELS_DIR}/${model_name}.part-"*)
    if [ "${#ordered_parts[@]}" -lt 1 ] || [ ! -e "${ordered_parts[0]}" ]; then
      continue
    fi

    cat "${ordered_parts[@]}" >"${SUBCREATOR_WHISPER_MODELS_CACHE_DIR}/${model_name}"
    copied_count=$((copied_count + 1))
  done

  if [ "${copied_count}" -gt 0 ]; then
    echo "Copied ${copied_count} bundled Whisper model(s) to ${SUBCREATOR_WHISPER_MODELS_CACHE_DIR}"
  fi
}

subcreator_json_escape() {
  # // Escape JSON string values safely for runtime-config generation.
  printf "%s" "$1" | sed 's/\\/\\\\/g; s/"/\\"/g'
}

subcreator_add_path_hint() {
  # // Keep path hints unique to avoid PATH duplication in CEP runtime.
  local hint="$1"
  if [ -z "${hint}" ]; then
    return 0
  fi

  case ":${SUBCREATOR_PATH_HINTS}:" in
    *":${hint}:"*)
      return 0
      ;;
  esac

  if [ -z "${SUBCREATOR_PATH_HINTS}" ]; then
    SUBCREATOR_PATH_HINTS="${hint}"
  else
    SUBCREATOR_PATH_HINTS="${SUBCREATOR_PATH_HINTS}:${hint}"
  fi
}

subcreator_resolve_python_executable_path() {
  # // Resolve the concrete interpreter path for the selected Python command.
  if [ -z "${SUBCREATOR_PYTHON_CMD}" ]; then
    return 1
  fi

  SUBCREATOR_PYTHON_PATH="$(${SUBCREATOR_PYTHON_CMD} -c 'import sys; print(sys.executable)' 2>/dev/null || true)"
  SUBCREATOR_PYTHON_PATH="$(printf "%s" "${SUBCREATOR_PYTHON_PATH}" | tr -d '\r' | sed -n '1p')"

  if [ -z "${SUBCREATOR_PYTHON_PATH}" ] && command -v "${SUBCREATOR_PYTHON_CMD}" >/dev/null 2>&1; then
    SUBCREATOR_PYTHON_PATH="$(command -v "${SUBCREATOR_PYTHON_CMD}")"
  fi

  if [ -n "${SUBCREATOR_PYTHON_PATH}" ]; then
    subcreator_add_path_hint "$(dirname "${SUBCREATOR_PYTHON_PATH}")"
    return 0
  fi

  return 1
}

subcreator_detect_whisper_path() {
  # // Detect the best whisper executable path for CEP runtime and host fallback usage.
  SUBCREATOR_WHISPER_PATH=""

  if [ -n "${SUBCREATOR_PYTHON_CMD}" ]; then
    local user_base=""
    user_base="$(${SUBCREATOR_PYTHON_CMD} -m site --user-base 2>/dev/null || true)"
    user_base="$(printf "%s" "${user_base}" | tr -d '\r' | sed -n '1p')"
    if [ -n "${user_base}" ]; then
      local user_whisper="${user_base}/bin/whisper"
      if [ -x "${user_whisper}" ]; then
        SUBCREATOR_WHISPER_PATH="${user_whisper}"
      fi
      subcreator_add_path_hint "${user_base}/bin"
    fi
  fi

  if [ -z "${SUBCREATOR_WHISPER_PATH}" ] && command -v whisper >/dev/null 2>&1; then
    SUBCREATOR_WHISPER_PATH="$(command -v whisper)"
  fi

  if [ -z "${SUBCREATOR_WHISPER_PATH}" ] && [ -n "${SUBCREATOR_PYTHON_PATH}" ]; then
    local sibling_whisper
    sibling_whisper="$(dirname "${SUBCREATOR_PYTHON_PATH}")/whisper"
    if [ -x "${sibling_whisper}" ]; then
      SUBCREATOR_WHISPER_PATH="${sibling_whisper}"
    fi
  fi

  if [ -n "${SUBCREATOR_WHISPER_PATH}" ]; then
    subcreator_add_path_hint "$(dirname "${SUBCREATOR_WHISPER_PATH}")"
    return 0
  fi

  return 1
}

subcreator_detect_ffmpeg_path() {
  # // Detect ffmpeg binary path so CEP commands can run without shell PATH assumptions.
  SUBCREATOR_FFMPEG_PATH=""
  if command -v ffmpeg >/dev/null 2>&1; then
    SUBCREATOR_FFMPEG_PATH="$(command -v ffmpeg)"
  elif [ -x "/opt/homebrew/bin/ffmpeg" ]; then
    SUBCREATOR_FFMPEG_PATH="/opt/homebrew/bin/ffmpeg"
  elif [ -x "/usr/local/bin/ffmpeg" ]; then
    SUBCREATOR_FFMPEG_PATH="/usr/local/bin/ffmpeg"
  fi

  if [ -n "${SUBCREATOR_FFMPEG_PATH}" ]; then
    subcreator_add_path_hint "$(dirname "${SUBCREATOR_FFMPEG_PATH}")"
    return 0
  fi

  return 1
}

subcreator_write_runtime_config() {
  # // Persist resolved runtime paths to a user-local config consumed by the extension.
  local generated_at
  generated_at="$(date -u +"%Y-%m-%dT%H:%M:%SZ")"

  subcreator_add_path_hint "/opt/homebrew/bin"
  subcreator_add_path_hint "/usr/local/bin"
  subcreator_add_path_hint "/usr/bin"
  subcreator_add_path_hint "/bin"

  mkdir -p "${SUBCREATOR_RUNTIME_DIR}"

  local path_hints_json=""
  local hint=""
  IFS=':' read -r -a hints <<<"${SUBCREATOR_PATH_HINTS}"
  for hint in "${hints[@]}"; do
    if [ -z "${hint}" ]; then
      continue
    fi
    if [ -n "${path_hints_json}" ]; then
      path_hints_json="${path_hints_json}, "
    fi
    path_hints_json="${path_hints_json}\"$(subcreator_json_escape "${hint}")\""
  done

  cat >"${SUBCREATOR_RUNTIME_FILE}" <<EOF
{
  "version": 1,
  "generatedBy": "subcreator_install_mac.sh",
  "generatedAtUtc": "${generated_at}",
  "pythonCommand": "$(subcreator_json_escape "${SUBCREATOR_PYTHON_CMD}")",
  "pythonPath": "$(subcreator_json_escape "${SUBCREATOR_PYTHON_PATH}")",
  "pythonVersion": "$(subcreator_json_escape "${SUBCREATOR_PYTHON_VERSION}")",
  "whisperPath": "$(subcreator_json_escape "${SUBCREATOR_WHISPER_PATH}")",
  "ffmpegPath": "$(subcreator_json_escape "${SUBCREATOR_FFMPEG_PATH}")",
  "pathHints": [${path_hints_json}]
}
EOF

  chmod 600 "${SUBCREATOR_RUNTIME_FILE}" || true

  echo "Runtime config written: ${SUBCREATOR_RUNTIME_FILE}"
  echo "  pythonPath=${SUBCREATOR_PYTHON_PATH:-<none>}"
  echo "  whisperPath=${SUBCREATOR_WHISPER_PATH:-<none>}"
  echo "  ffmpegPath=${SUBCREATOR_FFMPEG_PATH:-<none>}"
}

subcreator_configure_whisper_path() {
  # // Persist PATH update for Whisper CLI location produced by pip --user on macOS.
  local whisper_bin_path="${HOME}/Library/Python/${SUBCREATOR_PYTHON_VERSION}/bin"
  if [ ! -d "${whisper_bin_path}" ]; then
    return 1
  fi

  local updated=0
  if subcreator_append_path_to_profile "${HOME}/.zprofile" "${whisper_bin_path}"; then
    updated=1
    echo "Added Whisper PATH to ${HOME}/.zprofile"
  fi

  if subcreator_append_path_to_profile "${HOME}/.zshrc" "${whisper_bin_path}"; then
    updated=1
    echo "Added Whisper PATH to ${HOME}/.zshrc"
  fi

  if [ "${updated}" -eq 1 ]; then
    echo "Restart Terminal or run: source ~/.zprofile"
  fi

  return 0
}

subcreator_probe_python_version() {
  # // Return "<major>.<minor>" for a python executable, or empty when not callable.
  local candidate="$1"
  "${candidate}" -c 'import sys; print(f"{sys.version_info[0]}.{sys.version_info[1]}")' 2>/dev/null || true
}

subcreator_is_supported_python_version() {
  # // Whisper auto-install targets CPython 3.8 to 3.13 based on package metadata support.
  local version="$1"
  local major="${version%%.*}"
  local minor="${version#*.}"
  minor="${minor%%.*}"

  if ! [[ "${major}" =~ ^[0-9]+$ ]] || ! [[ "${minor}" =~ ^[0-9]+$ ]]; then
    return 1
  fi

  if [ "${major}" -ne 3 ]; then
    return 1
  fi

  if [ "${minor}" -lt 8 ] || [ "${minor}" -gt 13 ]; then
    return 1
  fi

  return 0
}

subcreator_supports_whisperx_version() {
  # // WhisperX currently requires CPython 3.10 to 3.13 for the corrected-align workflow.
  local version="$1"
  local major="${version%%.*}"
  local minor="${version#*.}"
  minor="${minor%%.*}"

  if ! [[ "${major}" =~ ^[0-9]+$ ]] || ! [[ "${minor}" =~ ^[0-9]+$ ]]; then
    return 1
  fi

  if [ "${major}" -ne 3 ]; then
    return 1
  fi

  if [ "${minor}" -lt 10 ] || [ "${minor}" -gt 13 ]; then
    return 1
  fi

  return 0
}

subcreator_validate_bundled_model_cache() {
  # // Confirm that the bundled base model is available from the cache location used by the panel.
  if [ -f "${SUBCREATOR_BUNDLED_MODELS_DIR}/base.pt" ] || [ -f "${SUBCREATOR_BUNDLED_MODELS_DIR}/base.pt.part-000" ]; then
    if [ -f "${SUBCREATOR_WHISPER_MODELS_CACHE_DIR}/base.pt" ]; then
      echo "Bundled Whisper base model available at ${SUBCREATOR_WHISPER_MODELS_CACHE_DIR}/base.pt"
    else
      echo "WARNING: bundled Whisper base model is missing from ${SUBCREATOR_WHISPER_MODELS_CACHE_DIR}"
    fi
  fi
}

subcreator_validate_whisper_install() {
  # // Confirm that the selected Python can import Whisper before the runtime config is written.
  if [ -z "${SUBCREATOR_PYTHON_CMD}" ]; then
    return 1
  fi

  if ${SUBCREATOR_PYTHON_CMD} -c 'import whisper' >/dev/null 2>&1; then
    echo "Whisper validation succeeded with ${SUBCREATOR_PYTHON_CMD}."
    return 0
  fi

  echo "Whisper validation failed with ${SUBCREATOR_PYTHON_CMD}."
  return 1
}

subcreator_validate_whisperx_install() {
  # // Confirm that the selected Python can import WhisperX before corrected align is advertised as ready.
  if [ -z "${SUBCREATOR_PYTHON_CMD}" ]; then
    return 1
  fi

  if ${SUBCREATOR_PYTHON_CMD} -c 'import whisperx' >/dev/null 2>&1; then
    echo "WhisperX validation succeeded with ${SUBCREATOR_PYTHON_CMD}."
    return 0
  fi

  echo "WhisperX validation failed with ${SUBCREATOR_PYTHON_CMD}."
  return 1
}

subcreator_select_python_cmd() {
  # // Prefer the same reliable Whisper/WhisperX Python versions as the Windows installer.
  local candidates=(
    "python3.11"
    "python3.12"
    "python3.10"
    "python3.13"
    "python3.9"
    "python3.8"
    "python3"
    "python"
  )

  local candidate=""
  local version=""
  for candidate in "${candidates[@]}"; do
    if ! command -v "${candidate}" >/dev/null 2>&1; then
      continue
    fi

    version="$(subcreator_probe_python_version "${candidate}")"
    if [ -z "${version}" ]; then
      continue
    fi

    if [ -n "${SUBCREATOR_PYTHON_SEEN}" ]; then
      SUBCREATOR_PYTHON_SEEN="${SUBCREATOR_PYTHON_SEEN}, "
    fi
    SUBCREATOR_PYTHON_SEEN="${SUBCREATOR_PYTHON_SEEN}${candidate}=${version}"

    if subcreator_is_supported_python_version "${version}"; then
      SUBCREATOR_PYTHON_CMD="${candidate}"
      SUBCREATOR_PYTHON_VERSION="${version}"
      return 0
    fi
  done

  return 1
}

# // Ensure built extension payload exists before copy.
if [ ! -d "${SUBCREATOR_SOURCE_DIR}" ]; then
  echo "Build missing: ${SUBCREATOR_SOURCE_DIR}"
  echo "Run: npm run subcreator:build"
  exit 1
fi

# // Create CEP extensions folder and copy payload atomically.
mkdir -p "$(dirname "${SUBCREATOR_DEST_DIR}")"
subcreator_backup_existing_templates
rm -rf "${SUBCREATOR_DEST_DIR}"
rm -rf "${SUBCREATOR_LEGACY_DEST_DIR}"
cp -R "${SUBCREATOR_SOURCE_DIR}" "${SUBCREATOR_DEST_DIR}"
subcreator_restore_existing_templates

echo "Sub Creator installed to ${SUBCREATOR_DEST_DIR}"
subcreator_enable_cep_debug_mode
subcreator_copy_bundled_whisper_models
subcreator_validate_bundled_model_cache

# // Discover supported Python runtime; when multiple versions exist we pick the newest supported one.
if ! subcreator_select_python_cmd; then
  if [ -z "${SUBCREATOR_PYTHON_SEEN}" ]; then
    echo "Whisper setup skipped: Python not found on this machine."
  else
    echo "Whisper setup skipped: no supported Python version found (need 3.8 to 3.13). Detected: ${SUBCREATOR_PYTHON_SEEN}"
  fi
  echo "Whisper and corrected align modes will remain unavailable in the panel until Python is installed."
else
  subcreator_resolve_python_executable_path || true
  echo "Installing Whisper with ${SUBCREATOR_PYTHON_CMD} (${SUBCREATOR_PYTHON_VERSION})..."

  # // Ensure pip is available, then install openai-whisper in user site-packages.
  if ! ${SUBCREATOR_PYTHON_CMD} -m pip --version >/dev/null 2>&1; then
    ${SUBCREATOR_PYTHON_CMD} -m ensurepip --upgrade >/dev/null 2>&1 || true
  fi

  if ${SUBCREATOR_PYTHON_CMD} -m pip install --user --upgrade openai-whisper; then
    echo "Whisper Python package installed successfully."
    subcreator_validate_whisper_install || true
    subcreator_configure_whisper_path || true
  else
    echo "Whisper package install failed. You can run manually:"
    echo "  ${SUBCREATOR_PYTHON_CMD} -m pip install --user --upgrade openai-whisper"
  fi

  # // Install WhisperX when the selected Python runtime is new enough for corrected transcript alignment.
  if subcreator_supports_whisperx_version "${SUBCREATOR_PYTHON_VERSION}"; then
    if ${SUBCREATOR_PYTHON_CMD} -m pip install --user --upgrade whisperx requests nltk certifi; then
      echo "WhisperX Python package installed successfully."
      subcreator_validate_whisperx_install || true
    else
      echo "WhisperX package install failed. You can run manually:"
      echo "  ${SUBCREATOR_PYTHON_CMD} -m pip install --user --upgrade whisperx requests nltk certifi"
    fi
  else
    echo "WhisperX setup skipped: Python ${SUBCREATOR_PYTHON_VERSION} detected (need 3.10 to 3.13 for corrected transcript align)."
  fi
fi

# // Install ffmpeg when Homebrew is present; otherwise keep install non-blocking.
if command -v ffmpeg >/dev/null 2>&1; then
  echo "ffmpeg already available."
elif command -v brew >/dev/null 2>&1; then
  echo "Installing ffmpeg via Homebrew..."
  if brew install ffmpeg; then
    echo "ffmpeg installed successfully."
  else
    echo "ffmpeg install failed. Install manually with: brew install ffmpeg"
  fi
else
  echo "ffmpeg not found and Homebrew unavailable. Install manually if Whisper transcription fails."
fi

subcreator_resolve_python_executable_path || true
subcreator_detect_whisper_path || true
subcreator_detect_ffmpeg_path || true
subcreator_write_runtime_config
echo "Restart Premiere Pro."
