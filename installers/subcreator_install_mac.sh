#!/usr/bin/env bash
set -euo pipefail

# // Resolve script and project directories reliably for macOS installation.
SUBCREATOR_SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SUBCREATOR_PROJECT_DIR="$(cd "${SUBCREATOR_SCRIPT_DIR}/.." && pwd)"
SUBCREATOR_SOURCE_DIR="${SUBCREATOR_PROJECT_DIR}/dist/com.cyrilg93.subcreator"
SUBCREATOR_DEST_DIR="${HOME}/Library/Application Support/Adobe/CEP/extensions/com.cyrilg93.subcreator"
SUBCREATOR_RUNTIME_DIR="${HOME}/Library/Application Support/SubCreator"
SUBCREATOR_RUNTIME_FILE="${SUBCREATOR_RUNTIME_DIR}/subcreator-runtime.json"
SUBCREATOR_PYTHON_CMD=""
SUBCREATOR_PYTHON_VERSION=""
SUBCREATOR_PYTHON_PATH=""
SUBCREATOR_PYTHON_SEEN=""
SUBCREATOR_WHISPER_PATH=""
SUBCREATOR_FFMPEG_PATH=""
SUBCREATOR_PATH_HINTS=""

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

subcreator_select_python_cmd() {
  # // Prefer explicit minor-version executables before generic python3/python aliases.
  local candidates=(
    "python3.13"
    "python3.12"
    "python3.11"
    "python3.10"
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
rm -rf "${SUBCREATOR_DEST_DIR}"
cp -R "${SUBCREATOR_SOURCE_DIR}" "${SUBCREATOR_DEST_DIR}"

echo "Sub Creator installed to ${SUBCREATOR_DEST_DIR}"
subcreator_enable_cep_debug_mode

# // Discover supported Python runtime; when multiple versions exist we pick the newest supported one.
if ! subcreator_select_python_cmd; then
  if [ -z "${SUBCREATOR_PYTHON_SEEN}" ]; then
    echo "Whisper setup skipped: Python not found on this machine."
  else
    echo "Whisper setup skipped: no supported Python version found (need 3.8 to 3.13). Detected: ${SUBCREATOR_PYTHON_SEEN}"
  fi
  echo "Whisper source will be hidden in the panel."
else
  subcreator_resolve_python_executable_path || true
  echo "Installing Whisper with ${SUBCREATOR_PYTHON_CMD} (${SUBCREATOR_PYTHON_VERSION})..."

  # // Ensure pip is available, then install openai-whisper in user site-packages.
  if ! ${SUBCREATOR_PYTHON_CMD} -m pip --version >/dev/null 2>&1; then
    ${SUBCREATOR_PYTHON_CMD} -m ensurepip --upgrade >/dev/null 2>&1 || true
  fi

  if ${SUBCREATOR_PYTHON_CMD} -m pip install --user --upgrade openai-whisper; then
    echo "Whisper Python package installed successfully."
    subcreator_configure_whisper_path || true
  else
    echo "Whisper package install failed. You can run manually:"
    echo "  ${SUBCREATOR_PYTHON_CMD} -m pip install --user --upgrade openai-whisper"
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
