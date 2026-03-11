#!/usr/bin/env bash
set -euo pipefail

# // Resolve script and project directories reliably for macOS installation.
SUBCREATOR_SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SUBCREATOR_PROJECT_DIR="$(cd "${SUBCREATOR_SCRIPT_DIR}/.." && pwd)"
SUBCREATOR_SOURCE_DIR="${SUBCREATOR_PROJECT_DIR}/dist/com.cyrilg93.subcreator"
SUBCREATOR_DEST_DIR="${HOME}/Library/Application Support/Adobe/CEP/extensions/com.cyrilg93.subcreator"
SUBCREATOR_PYTHON_CMD=""

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

# // Discover local Python runtime; if absent we skip Whisper install as requested.
if command -v python3 >/dev/null 2>&1; then
  SUBCREATOR_PYTHON_CMD="python3"
elif command -v python >/dev/null 2>&1; then
  SUBCREATOR_PYTHON_CMD="python"
fi

if [ -z "${SUBCREATOR_PYTHON_CMD}" ]; then
  echo "Whisper setup skipped: Python not found on this machine."
  echo "Whisper source will be hidden in the panel."
  echo "If needed, enable CEP debug mode and restart Premiere Pro."
  exit 0
fi

# // Read Python major/minor to avoid unsupported 3.14+ auto-install.
SUBCREATOR_PYTHON_VERSION="$(${SUBCREATOR_PYTHON_CMD} -c 'import sys; print(f"{sys.version_info[0]}.{sys.version_info[1]}")' 2>/dev/null || true)"
SUBCREATOR_PYTHON_MAJOR="${SUBCREATOR_PYTHON_VERSION%%.*}"
SUBCREATOR_PYTHON_MINOR="${SUBCREATOR_PYTHON_VERSION#*.}"
SUBCREATOR_PYTHON_MINOR="${SUBCREATOR_PYTHON_MINOR%%.*}"

if ! [[ "${SUBCREATOR_PYTHON_MAJOR}" =~ ^[0-9]+$ ]] || ! [[ "${SUBCREATOR_PYTHON_MINOR}" =~ ^[0-9]+$ ]]; then
  echo "Whisper setup skipped: unable to parse Python version from '${SUBCREATOR_PYTHON_VERSION}'."
  echo "If needed, enable CEP debug mode and restart Premiere Pro."
  exit 0
fi

if [ "${SUBCREATOR_PYTHON_MAJOR}" -gt 3 ] || { [ "${SUBCREATOR_PYTHON_MAJOR}" -eq 3 ] && [ "${SUBCREATOR_PYTHON_MINOR}" -ge 14 ]; }; then
  echo "Whisper setup skipped: Python ${SUBCREATOR_PYTHON_VERSION} detected (openai-whisper currently targets Python <= 3.13)."
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
