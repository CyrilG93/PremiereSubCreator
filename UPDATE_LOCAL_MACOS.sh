#!/bin/bash
set -u

# // Build and copy the local CEP extension without rebuilding the macOS installer.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
"${SCRIPT_DIR}/scripts/subcreator-update-local-macos.sh" "$@"
EXIT_CODE=$?

if [ "${EXIT_CODE}" -ne 0 ]; then
  echo
  echo "Local update failed with code ${EXIT_CODE}."
  exit "${EXIT_CODE}"
fi

echo
echo "Local update complete. Restart Premiere Pro to reload the panel."
