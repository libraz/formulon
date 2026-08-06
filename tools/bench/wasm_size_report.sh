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
#                                   [--brotli-ceiling-bytes N]
#                                   [--brotli-soft-ceiling-bytes N]
#
# PATH defaults to build-wasm/formulon.wasm.
#
# Defaults (per CLAUDE.md "WASM Size Policy"):
#   --ceiling-bytes             3145728   (3.00 MiB hard ceiling)
#   --soft-ceiling-bytes        2621440   (2.50 MiB stretch goal)
#   --brotli-ceiling-bytes      786432    (768 KiB hard ceiling)
#   --brotli-soft-ceiling-bytes 655360    (640 KiB stretch goal)
#
# Brotli wire size is the binding constraint in practice, so it is gated on
# equal footing with the uncompressed size rather than merely reported. When
# `brotli` is absent from PATH the Brotli check is skipped, not failed --
# a missing tool must not turn into a phantom size regression.
#
# Exit codes:
#   0 -- both sizes at or under their hard ceilings (warns on stderr if either
#        exceeds its soft ceiling)
#   1 -- a hard ceiling exceeded
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
       [--brotli-ceiling-bytes N] [--brotli-soft-ceiling-bytes N]

Report the size of a Formulon .wasm artifact and check it against the
configured size ceilings.

Arguments:
  PATH                            Path to the .wasm file (default: build-wasm/formulon.wasm)
  --json                          Emit a single-object JSON document instead of text.
  --ceiling-bytes N               Hard ceiling in bytes (default 3145728 = 3.00 MiB).
  --soft-ceiling-bytes N          Soft ceiling in bytes (default 2621440 = 2.50 MiB).
  --brotli-ceiling-bytes N        Brotli hard ceiling in bytes (default 786432 = 768 KiB).
  --brotli-soft-ceiling-bytes N   Brotli soft ceiling in bytes (default 655360 = 640 KiB).
  -h, --help                      Show this help.

Exit codes: 0 ok, 1 hard ceiling exceeded, 2 artifact missing, 3 bad args.
EOF
}

# Defaults.
WASM_PATH=""
EMIT_JSON=0
HARD_CEILING=3145728
SOFT_CEILING=2621440
BROTLI_HARD_CEILING=786432
BROTLI_SOFT_CEILING=655360

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
    --brotli-ceiling-bytes)
      if [ $# -lt 2 ]; then
        printf '%s: --brotli-ceiling-bytes requires a value\n' "$PROG" >&2
        exit 3
      fi
      BROTLI_HARD_CEILING="$2"
      shift 2
      ;;
    --brotli-ceiling-bytes=*)
      BROTLI_HARD_CEILING="${1#--brotli-ceiling-bytes=}"
      shift
      ;;
    --brotli-soft-ceiling-bytes)
      if [ $# -lt 2 ]; then
        printf '%s: --brotli-soft-ceiling-bytes requires a value\n' "$PROG" >&2
        exit 3
      fi
      BROTLI_SOFT_CEILING="$2"
      shift 2
      ;;
    --brotli-soft-ceiling-bytes=*)
      BROTLI_SOFT_CEILING="${1#--brotli-soft-ceiling-bytes=}"
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
validate_int "--brotli-ceiling-bytes" "$BROTLI_HARD_CEILING"
validate_int "--brotli-soft-ceiling-bytes" "$BROTLI_SOFT_CEILING"

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

# Determine per-measure status: ok / soft-warn / over-hard-ceiling.
UNCOMPRESSED_STATUS="ok"
if [ "$UNCOMPRESSED_BYTES" -gt "$HARD_CEILING" ]; then
  UNCOMPRESSED_STATUS="over-hard-ceiling"
elif [ "$UNCOMPRESSED_BYTES" -gt "$SOFT_CEILING" ]; then
  UNCOMPRESSED_STATUS="over-soft-ceiling"
fi

BROTLI_STATUS="skipped"
if [ -n "$BROTLI_BYTES" ]; then
  BROTLI_STATUS="ok"
  if [ "$BROTLI_BYTES" -gt "$BROTLI_HARD_CEILING" ]; then
    BROTLI_STATUS="over-hard-ceiling"
  elif [ "$BROTLI_BYTES" -gt "$BROTLI_SOFT_CEILING" ]; then
    BROTLI_STATUS="over-soft-ceiling"
  fi
fi

# The overall status is the worse of the two, so a Brotli overage fails the
# gate even while the uncompressed size still has headroom.
STATUS="ok"
EXIT_CODE=0
if [ "$UNCOMPRESSED_STATUS" = "over-hard-ceiling" ] || [ "$BROTLI_STATUS" = "over-hard-ceiling" ]; then
  STATUS="over-hard-ceiling"
  EXIT_CODE=1
elif [ "$UNCOMPRESSED_STATUS" = "over-soft-ceiling" ] || [ "$BROTLI_STATUS" = "over-soft-ceiling" ]; then
  STATUS="over-soft-ceiling"
fi

# Compute headroom against the hard ceiling. May be negative.
HEADROOM_BYTES="$(awk -v h="$HARD_CEILING" -v u="$UNCOMPRESSED_BYTES" \
  'BEGIN { printf "%d", h - u }')"
HEADROOM_PCT="$(awk -v h="$HARD_CEILING" -v u="$UNCOMPRESSED_BYTES" \
  'BEGIN { if (h <= 0) { print "0.0"; exit } printf "%.1f", (h - u) * 100.0 / h }')"
if [ -n "$BROTLI_BYTES" ]; then
  BROTLI_HEADROOM_BYTES="$(awk -v h="$BROTLI_HARD_CEILING" -v u="$BROTLI_BYTES" \
    'BEGIN { printf "%d", h - u }')"
  BROTLI_HEADROOM_PCT="$(awk -v h="$BROTLI_HARD_CEILING" -v u="$BROTLI_BYTES" \
    'BEGIN { if (h <= 0) { print "0.0"; exit } printf "%.1f", (h - u) * 100.0 / h }')"
else
  BROTLI_HEADROOM_BYTES=""
  BROTLI_HEADROOM_PCT=""
fi

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
  printf '"brotli_hard_ceiling_bytes":%s,' "$BROTLI_HARD_CEILING"
  printf '"brotli_soft_ceiling_bytes":%s,' "$BROTLI_SOFT_CEILING"
  printf '"headroom_bytes":%s,' "$HEADROOM_BYTES"
  printf '"headroom_pct":%s,' "$HEADROOM_PCT"
  if [ -n "$BROTLI_HEADROOM_BYTES" ]; then
    printf '"brotli_headroom_bytes":%s,' "$BROTLI_HEADROOM_BYTES"
    printf '"brotli_headroom_pct":%s,' "$BROTLI_HEADROOM_PCT"
  else
    printf '"brotli_headroom_bytes":null,'
    printf '"brotli_headroom_pct":null,'
  fi
  printf '"uncompressed_status":"%s",' "$UNCOMPRESSED_STATUS"
  printf '"brotli_status":"%s",' "$BROTLI_STATUS"
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
  printf '  br-ceiling:   %s bytes (%s KiB)   [hard ceiling]\n' \
    "$BROTLI_HARD_CEILING" "$(kib "$BROTLI_HARD_CEILING")"
  printf '  br-soft:      %s bytes (%s KiB)   [stretch goal]\n' \
    "$BROTLI_SOFT_CEILING" "$(kib "$BROTLI_SOFT_CEILING")"
  if [ -n "$BROTLI_HEADROOM_BYTES" ]; then
    printf '  br-headroom:  %s bytes (%s%%)\n' "$BROTLI_HEADROOM_BYTES" "$BROTLI_HEADROOM_PCT"
  else
    printf '  br-headroom:  (skipped)\n'
  fi
  printf '  status:       %s (uncompressed: %s, brotli: %s)\n' \
    "$STATUS" "$UNCOMPRESSED_STATUS" "$BROTLI_STATUS"
fi

if [ "$UNCOMPRESSED_STATUS" = "over-soft-ceiling" ]; then
  printf '%s: warning: %s bytes exceeds soft ceiling %s bytes\n' \
    "$PROG" "$UNCOMPRESSED_BYTES" "$SOFT_CEILING" >&2
fi
if [ "$UNCOMPRESSED_STATUS" = "over-hard-ceiling" ]; then
  printf '%s: error: %s bytes exceeds hard ceiling %s bytes\n' \
    "$PROG" "$UNCOMPRESSED_BYTES" "$HARD_CEILING" >&2
fi
if [ "$BROTLI_STATUS" = "over-soft-ceiling" ]; then
  printf '%s: warning: brotli %s bytes exceeds soft ceiling %s bytes\n' \
    "$PROG" "$BROTLI_BYTES" "$BROTLI_SOFT_CEILING" >&2
fi
if [ "$BROTLI_STATUS" = "over-hard-ceiling" ]; then
  printf '%s: error: brotli %s bytes exceeds hard ceiling %s bytes\n' \
    "$PROG" "$BROTLI_BYTES" "$BROTLI_HARD_CEILING" >&2
fi

exit "$EXIT_CODE"
