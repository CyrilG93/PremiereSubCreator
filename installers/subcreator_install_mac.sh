#!/usr/bin/env bash
set -euo pipefail

# // Resolve script and project directories reliably for macOS installation.
SUBCREATOR_SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SUBCREATOR_PROJECT_DIR="$(cd "${SUBCREATOR_SCRIPT_DIR}/.." && pwd)"
SUBCREATOR_SOURCE_DIR="${SUBCREATOR_PROJECT_DIR}/dist/com.cyrilg93.subcreator"
SUBCREATOR_DEST_DIR="${HOME}/Library/Application Support/Adobe/CEP/extensions/com.cyrilg93.subcreator"

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
echo "If needed, enable CEP debug mode and restart Premiere Pro."
