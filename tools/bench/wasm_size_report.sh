#!/bin/sh
# wasm_size_report.sh
#
# Report the size of a built `.wasm` artifact in human-readable or JSON form
# and gate against a configurable size ceiling.
#
# This is the standalone counterpart to cmake/FormulonWasmSizeReport.cmake;
# the CMake reporter runs as a POST_BUILD step inside the WASM build, while
# this script is intended for CI workflows, PR comment scrapers, and the
# `make size-check` / `/size-check` slash command.
#
# Usage:
#   tools/bench/wasm_size_report.sh [PATH] [--json]
#                                   [--ceiling-bytes N] [--soft-ceiling-bytes N]
#
# PATH defaults to build-wasm/formulon.wasm.
#
# Defaults (per CLAUDE.md "WASM Size Policy"):
#   --ceiling-bytes      3145728   (3.00 MiB hard ceiling)
#   --soft-ceiling-bytes 2621440   (2.50 MiB stretch goal)
#
# Exit codes:
#   0 -- size at or under the hard ceiling (warns on stderr if it exceeds the
#        soft ceiling)
#   1 -- hard ceiling exceeded
#   2 -- artifact missing
#   3 -- argument error
#
# POSIX sh-only: no bashisms (no [[ ]], no arrays, no $(( float ))). Floating
# point math is delegated to awk for portability across BSD / GNU / busybox.

set -eu

PROG="$(basename "$0")"

usage() {
  cat <<EOF
Usage: ${PROG} [PATH] [--json] [--ceiling-bytes N] [--soft-ceiling-bytes N]

Report the size of a Formulon .wasm artifact and check it against the
configured size ceilings.

Arguments:
  PATH                       Path to the .wasm file (default: build-wasm/formulon.wasm)
  --json                     Emit a single-object JSON document instead of text.
  --ceiling-bytes N          Hard ceiling in bytes (default 3145728 = 3.00 MiB).
  --soft-ceiling-bytes N     Soft ceiling in bytes (default 2621440 = 2.50 MiB).
  -h, --help                 Show this help.

Exit codes: 0 ok, 1 hard ceiling exceeded, 2 artifact missing, 3 bad args.
EOF
}

# Defaults.
WASM_PATH=""
EMIT_JSON=0
HARD_CEILING=3145728
SOFT_CEILING=2621440

# Argument parsing (POSIX sh, no getopts long-opt support).
while [ $# -gt 0 ]; do
  case "$1" in
    --json)
      EMIT_JSON=1
      shift
      ;;
    --ceiling-bytes)
      if [ $# -lt 2 ]; then
        printf '%s: --ceiling-bytes requires a value\n' "$PROG" >&2
        exit 3
      fi
      HARD_CEILING="$2"
      shift 2
      ;;
    --ceiling-bytes=*)
      HARD_CEILING="${1#--ceiling-bytes=}"
      shift
      ;;
    --soft-ceiling-bytes)
      if [ $# -lt 2 ]; then
        printf '%s: --soft-ceiling-bytes requires a value\n' "$PROG" >&2
        exit 3
      fi
      SOFT_CEILING="$2"
      shift 2
      ;;
    --soft-ceiling-bytes=*)
      SOFT_CEILING="${1#--soft-ceiling-bytes=}"
      shift
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    --*)
      printf '%s: unknown option: %s\n' "$PROG" "$1" >&2
      usage >&2
      exit 3
      ;;
    *)
      if [ -z "$WASM_PATH" ]; then
        WASM_PATH="$1"
        shift
      else
        printf '%s: unexpected positional argument: %s\n' "$PROG" "$1" >&2
        exit 3
      fi
      ;;
  esac
done

if [ -z "$WASM_PATH" ]; then
  WASM_PATH="build-wasm/formulon.wasm"
fi

# Validate the integer arguments. awk used because POSIX sh `expr` rejects
# leading zeros and negatives inconsistently across platforms.
validate_int() {
  _label="$1"
  _value="$2"
  case "$_value" in
    ''|*[!0-9]*)
      printf '%s: %s must be a non-negative integer (got %s)\n' \
        "$PROG" "$_label" "$_value" >&2
      exit 3
      ;;
  esac
}
validate_int "--ceiling-bytes" "$HARD_CEILING"
validate_int "--soft-ceiling-bytes" "$SOFT_CEILING"

if [ ! -f "$WASM_PATH" ]; then
  if [ "$EMIT_JSON" -eq 1 ]; then
    # Emit a structured error so CI scrapers can still parse it.
    printf '{"path":"%s","status":"missing","error":"artifact not found"}\n' \
      "$WASM_PATH"
  else
    printf '%s: artifact not found: %s\n' "$PROG" "$WASM_PATH" >&2
  fi
  exit 2
fi

# Uncompressed bytes via wc -c (portable).
UNCOMPRESSED_BYTES="$(wc -c < "$WASM_PATH" | tr -d ' \t\n')"
validate_int "uncompressed-size" "$UNCOMPRESSED_BYTES"

# Brotli (best-effort). brotli writes to <input>.br by default; we use
# stdout to avoid filesystem side effects in CI scratch dirs.
BROTLI_BYTES=""
BROTLI_TOOL_NOTE="missing - install for Brotli reporting"
if command -v brotli >/dev/null 2>&1; then
  BROTLI_TOOL_NOTE="brotli"
  # `brotli --quality=11` is the default but spell it out for determinism.
  # `--stdout` plus `--force` keeps us out of the filesystem.
  if brotli --quality=11 --force --stdout < "$WASM_PATH" > "${WASM_PATH}.br.tmp" 2>/dev/null; then
    BROTLI_BYTES="$(wc -c < "${WASM_PATH}.br.tmp" | tr -d ' \t\n')"
    rm -f "${WASM_PATH}.br.tmp"
  else
    rm -f "${WASM_PATH}.br.tmp"
    BROTLI_TOOL_NOTE="brotli (compression failed)"
  fi
fi

# Determine status: ok / soft-warn / over-hard-ceiling.
STATUS="ok"
EXIT_CODE=0
if [ "$UNCOMPRESSED_BYTES" -gt "$HARD_CEILING" ]; then
  STATUS="over-hard-ceiling"
  EXIT_CODE=1
elif [ "$UNCOMPRESSED_BYTES" -gt "$SOFT_CEILING" ]; then
  STATUS="over-soft-ceiling"
  EXIT_CODE=0
fi

# Compute headroom against the hard ceiling. May be negative.
HEADROOM_BYTES="$(awk -v h="$HARD_CEILING" -v u="$UNCOMPRESSED_BYTES" \
  'BEGIN { printf "%d", h - u }')"
HEADROOM_PCT="$(awk -v h="$HARD_CEILING" -v u="$UNCOMPRESSED_BYTES" \
  'BEGIN { if (h <= 0) { print "0.0"; exit } printf "%.1f", (h - u) * 100.0 / h }')"

# Pretty MiB / KiB. awk handles the floats.
mib() { awk -v b="$1" 'BEGIN { printf "%.2f", b / 1048576.0 }'; }
kib() { awk -v b="$1" 'BEGIN { printf "%.0f", b / 1024.0 }'; }

WASM_NAME="$(basename "$WASM_PATH")"

if [ "$EMIT_JSON" -eq 1 ]; then
  # Single line; no trailing comma; field order matches the human report.
  printf '{'
  printf '"path":"%s",' "$WASM_PATH"
  printf '"name":"%s",' "$WASM_NAME"
  printf '"uncompressed_bytes":%s,' "$UNCOMPRESSED_BYTES"
  if [ -n "$BROTLI_BYTES" ]; then
    printf '"brotli_bytes":%s,' "$BROTLI_BYTES"
  else
    printf '"brotli_bytes":null,'
  fi
  printf '"brotli_tool":"%s",' "$BROTLI_TOOL_NOTE"
  printf '"hard_ceiling_bytes":%s,' "$HARD_CEILING"
  printf '"soft_ceiling_bytes":%s,' "$SOFT_CEILING"
  printf '"headroom_bytes":%s,' "$HEADROOM_BYTES"
  printf '"headroom_pct":%s,' "$HEADROOM_PCT"
  printf '"status":"%s"' "$STATUS"
  printf '}\n'
else
  printf '%s\n' "$WASM_NAME"
  printf '  uncompressed: %s bytes (%s MiB)\n' "$UNCOMPRESSED_BYTES" "$(mib "$UNCOMPRESSED_BYTES")"
  if [ -n "$BROTLI_BYTES" ]; then
    printf '  brotli:       %s bytes (%s KiB)   [tool: %s]\n' \
      "$BROTLI_BYTES" "$(kib "$BROTLI_BYTES")" "$BROTLI_TOOL_NOTE"
  else
    printf '  brotli:       (skipped)                  [tool: %s]\n' "$BROTLI_TOOL_NOTE"
  fi
  printf '  ceiling:      %s bytes (%s MiB)   [hard ceiling]\n' \
    "$HARD_CEILING" "$(mib "$HARD_CEILING")"
  printf '  soft-ceiling: %s bytes (%s MiB)   [stretch goal]\n' \
    "$SOFT_CEILING" "$(mib "$SOFT_CEILING")"
  printf '  headroom:     %s bytes (%s%%)\n' "$HEADROOM_BYTES" "$HEADROOM_PCT"
  printf '  status:       %s\n' "$STATUS"
fi

if [ "$STATUS" = "over-soft-ceiling" ]; then
  printf '%s: warning: %s bytes exceeds soft ceiling %s bytes\n' \
    "$PROG" "$UNCOMPRESSED_BYTES" "$SOFT_CEILING" >&2
fi
if [ "$STATUS" = "over-hard-ceiling" ]; then
  printf '%s: error: %s bytes exceeds hard ceiling %s bytes\n' \
    "$PROG" "$UNCOMPRESSED_BYTES" "$HARD_CEILING" >&2
fi

exit "$EXIT_CODE"
