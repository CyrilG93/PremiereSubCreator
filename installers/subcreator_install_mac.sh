#!/usr/bin/env bash
set -euo pipefail

# // Resolve script and project directories reliably for macOS installation.
SUBCREATOR_SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SUBCREATOR_PROJECT_DIR="$(cd "${SUBCREATOR_SCRIPT_DIR}/.." && pwd)"
SUBCREATOR_SOURCE_DIR="${SUBCREATOR_PROJECT_DIR}/dist/com.cyrilg93.subcreator"
SUBCREATOR_DEST_DIR="${HOME}/Library/Application Support/Adobe/CEP/extensions/com.cyrilg93.subcreator"
SUBCREATOR_PYTHON_CMD=""
SUBCREATOR_PYTHON_VERSION=""
SUBCREATOR_PYTHON_SEEN=""

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

# // Discover supported Python runtime; when multiple versions exist we pick the newest supported one.
if ! subcreator_select_python_cmd; then
  if [ -z "${SUBCREATOR_PYTHON_SEEN}" ]; then
    echo "Whisper setup skipped: Python not found on this machine."
  else
    echo "Whisper setup skipped: no supported Python version found (need 3.8 to 3.13). Detected: ${SUBCREATOR_PYTHON_SEEN}"
  fi
  echo "Whisper source will be hidden in the panel."
  echo "If needed, enable CEP debug mode and restart Premiere Pro."
  exit 0
fi

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

echo "If needed, enable CEP debug mode and restart Premiere Pro."
