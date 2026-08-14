#!/bin/sh
# emsdk_pin.sh
#
# Print the pinned Emscripten SDK version held in tools/wasm/emsdk-version.txt.
#
# Having one reader keeps the parsing rule (first non-comment, non-blank word)
# in a single place, so the workflows and the size reporter can never disagree
# about what the pin says.
#
# Usage:
#   tools/wasm/emsdk_pin.sh
#
# Exit codes:
#   0 -- version printed on stdout
#   1 -- version file missing or empty

set -eu

PROG="$(basename "$0")"
PIN_FILE="$(dirname "$0")/emsdk-version.txt"

if [ ! -f "$PIN_FILE" ]; then
  printf '%s: pin file not found: %s\n' "$PROG" "$PIN_FILE" >&2
  exit 1
fi

VERSION="$(awk '!/^[[:space:]]*#/ && NF { print $1; exit }' "$PIN_FILE")"

if [ -z "$VERSION" ]; then
  printf '%s: no version line in %s\n' "$PROG" "$PIN_FILE" >&2
  exit 1
fi

printf '%s\n' "$VERSION"
