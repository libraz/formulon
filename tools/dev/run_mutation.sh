#!/bin/sh
# Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
#
# run_mutation.sh
#
# Local mutation-testing diagnostic. Builds Formulon with the
# `mutation` CMake preset (clang + LTO + embedded bitcode, the
# combination mull-runner-cxx requires), then mutates a focused
# subset of source files and runs the unit-test binary against
# every surviving mutant. Prints a single mutation-score summary
# and the path to the full IDE-reporter log.
#
# This is a *local* diagnostic, not a CI gate. CI is verification,
# not the development loop, and a fixed mutation-score threshold is
# the same anti-pattern as a fixed coverage threshold: it drifts
# under benign refactors and incentivises tests that kill mutants
# without actually asserting anything useful. The default mode here
# prints the score and exits 0.
#
# A soft target of ~70% on the configured filter set (hot helpers
# under src/eval/) is reasonable when triaging a "are these helpers
# actually exercised?" question; pass FORMULON_MUT_STRICT=1 for
# ad-hoc local enforcement. CI never sets it.
#
# Usage:
#   tools/dev/run_mutation.sh
#   FORMULON_MUT_STRICT=1 tools/dev/run_mutation.sh
#   FORMULON_MUT_FILTER=src/eval/coerce.cpp tools/dev/run_mutation.sh
#
# Environment overrides:
#   FORMULON_MUT_BUILD_DIR  Build directory (default: build-mutation)
#   FORMULON_MUT_TIMEOUT    Per-mutant test timeout, seconds
#                           (default: 120; mull's --timeout takes ms)
#   FORMULON_MUT_STRICT     If non-empty, exit 1 when score < 70
#                           (default: report only)
#   FORMULON_MUT_TARGET     Test binary mull mutates against
#                           (default: formulon_unit_tests)
#   FORMULON_MUT_FILTER     Comma-separated path globs limiting the
#                           mutation scope. Each entry becomes a
#                           --include-path=<glob> flag for mull.
#                           Default: src/eval/coerce.cpp,
#                                    src/eval/criteria.cpp,
#                                    src/eval/text_ops.cpp
#   MULL_RUNNER             Mull runner binary (default: mull-runner-cxx)
#   CMAKE                   cmake binary (default: cmake)
#   CTEST                   ctest binary (default: ctest)
#
# Exit codes:
#   0 -- mutation report rendered (regardless of score, unless
#        FORMULON_MUT_STRICT is set)
#   1 -- FORMULON_MUT_STRICT was set and the score was below the
#        soft target
#   2 -- configure / build / tooling-missing failure

set -eu

PROG="$(basename "$0")"

BUILD_DIR="${FORMULON_MUT_BUILD_DIR:-build-mutation}"
TIMEOUT_SEC="${FORMULON_MUT_TIMEOUT:-120}"
STRICT_FLAG="${FORMULON_MUT_STRICT:-}"
TARGET_BIN="${FORMULON_MUT_TARGET:-formulon_unit_tests}"
FILTER_RAW="${FORMULON_MUT_FILTER:-src/eval/coerce.cpp,src/eval/criteria.cpp,src/eval/text_ops.cpp}"
MULL_RUNNER="${MULL_RUNNER:-mull-runner-cxx}"
CMAKE="${CMAKE:-cmake}"
CTEST="${CTEST:-ctest}"

# Soft target used by --strict mode. Not a CI gate.
SOFT_TARGET_PCT=70

# Find the repo root: assume this script lives at <repo>/tools/dev/run_mutation.sh.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../.." && pwd)"
cd "${REPO_ROOT}"

LOG_DIR="${REPO_ROOT}/${BUILD_DIR}/_mut_logs"
mkdir -p "${LOG_DIR}"
CONFIG_LOG="${LOG_DIR}/configure.log"
BUILD_LOG="${LOG_DIR}/build.log"
RUN_LOG="${LOG_DIR}/mutation.log"

# Tooling preflight. mull is clang-only -- the gcc path doesn't
# emit the LLVM bitcode mull-runner-cxx walks.
if ! command -v "${MULL_RUNNER}" >/dev/null 2>&1; then
  echo "${PROG}: '${MULL_RUNNER}' not found in PATH; install mull (brew: mull-project/mull/mull, apt: mull on Linux). See https://mull.readthedocs.io/ for setup." >&2
  exit 2
fi
if ! command -v clang++ >/dev/null 2>&1; then
  echo "${PROG}: 'clang++' not found in PATH; mull-runner-cxx requires a clang toolchain (brew: llvm, apt: clang)." >&2
  exit 2
fi

# Translate FORMULON_MUT_FILTER (comma-separated globs) into a list
# of --include-path=<glob> flags. POSIX shell, no arrays: rebuild
# positional parameters and re-quote on use. An empty filter means
# "no include filter" -- mull will mutate everything it sees.
INCLUDE_FLAGS=""
if [ -n "${FILTER_RAW}" ]; then
  # Split on comma using IFS; preserve original IFS afterwards.
  OLD_IFS="${IFS}"
  IFS=','
  set -- ${FILTER_RAW}
  IFS="${OLD_IFS}"
  for entry in "$@"; do
    # Trim leading/trailing whitespace via parameter expansion +
    # awk; awk is POSIX and avoids `sed -E`. dash has no
    # ${var# / %} repeated-pattern, so awk is simpler.
    trimmed="$(printf '%s' "${entry}" | awk '{$1=$1};1')"
    if [ -n "${trimmed}" ]; then
      INCLUDE_FLAGS="${INCLUDE_FLAGS} --include-path=${trimmed}"
    fi
  done
fi

echo "${PROG}: mutation build dir: ${BUILD_DIR}"
echo "${PROG}: target binary: ${TARGET_BIN}"
echo "${PROG}: include filter: ${FILTER_RAW:-<none>}"
echo "${PROG}: configuring (preset=mutation)..."

# The `mutation` configure preset is declared in CMakePresets.json
# and wires up clang + -O0 -g -fembed-bitcode -flto=full so that
# mull-runner-cxx can recover LLVM bitcode from the test binary.
if ! "${CMAKE}" --preset mutation > "${CONFIG_LOG}" 2>&1; then
  echo "${PROG}: cmake configure failed; see ${CONFIG_LOG}" >&2
  tail -50 "${CONFIG_LOG}" >&2 || true
  exit 2
fi

echo "${PROG}: building target ${TARGET_BIN}..."
if ! "${CMAKE}" --build "${BUILD_DIR}" --parallel --target "${TARGET_BIN}" > "${BUILD_LOG}" 2>&1; then
  echo "${PROG}: build failed; see ${BUILD_LOG}" >&2
  tail -100 "${BUILD_LOG}" >&2 || true
  exit 2
fi

# mull-runner-cxx expects the binary path at the end of its argv.
# The test binary is normally produced under build/bin/, but some
# CMake configs write to build/tests/ or directly into BUILD_DIR.
# Probe the common locations and pick the first hit.
BIN_PATH=""
for candidate in \
    "${BUILD_DIR}/bin/${TARGET_BIN}" \
    "${BUILD_DIR}/tests/unit/${TARGET_BIN}" \
    "${BUILD_DIR}/tests/${TARGET_BIN}" \
    "${BUILD_DIR}/${TARGET_BIN}"; do
  if [ -x "${candidate}" ]; then
    BIN_PATH="${candidate}"
    break
  fi
done

if [ -z "${BIN_PATH}" ]; then
  echo "${PROG}: built ${TARGET_BIN} but could not locate the executable under ${BUILD_DIR}; see ${BUILD_LOG}" >&2
  exit 2
fi

# mull's --timeout is in milliseconds.
TIMEOUT_MS=$((TIMEOUT_SEC * 1000))

echo "${PROG}: running mull on ${BIN_PATH} (per-mutant timeout ${TIMEOUT_SEC}s)..."

# mull may exit non-zero when surviving mutants exist -- that's a
# finding, not a tooling failure. Capture rc and continue so we
# can still emit the score summary.
set +e
# shellcheck disable=SC2086
"${MULL_RUNNER}" \
    --report-name=formulon-mutation \
    --report-dir="${LOG_DIR}" \
    --reporters=IDE \
    --timeout="${TIMEOUT_MS}" \
    ${INCLUDE_FLAGS} \
    "${BIN_PATH}" \
    > "${RUN_LOG}" 2>&1
MULL_RC=$?
set -e

# Always echo the run log so the user can see which mutants
# survived. The IDE reporter writes one line per surviving mutant
# plus a final "Mutation Score: NN.N%" summary.
cat "${RUN_LOG}"

# Extract the mutation score. mull's IDE reporter prints something
# like "Mutation Score: 73.2%". Use awk (POSIX) instead of sed -E.
SCORE_PCT="$(awk '
  /Mutation Score:/ {
    for (i = 1; i <= NF; i++) {
      if ($i ~ /%$/) {
        gsub("%", "", $i)
        print $i
        exit
      }
    }
  }
' "${RUN_LOG}")"

echo ""
if [ -n "${SCORE_PCT}" ]; then
  echo "${PROG}: mutation score: ${SCORE_PCT}%"
else
  echo "${PROG}: mutation score: <not reported by mull>"
fi
echo "${PROG}: full log: ${RUN_LOG}"
echo "${PROG}: configure log: ${CONFIG_LOG}"
echo "${PROG}: build log: ${BUILD_LOG}"

# If mull failed for a reason other than surviving mutants (e.g.
# binary it couldn't open, missing symbols), surface that. We
# treat the absence of a parsed score as a tooling failure when
# the rc is also non-zero.
if [ -z "${SCORE_PCT}" ] && [ "${MULL_RC}" -ne 0 ]; then
  echo "${PROG}: mull exited rc=${MULL_RC} and produced no score; treating as tooling failure" >&2
  exit 2
fi

# Strict-mode threshold check. POSIX has no float arithmetic, so
# round to an integer percentage for the comparison; awk does the
# truncation cleanly.
if [ -n "${STRICT_FLAG}" ] && [ -n "${SCORE_PCT}" ]; then
  SCORE_INT="$(awk -v s="${SCORE_PCT}" 'BEGIN { printf "%d", s + 0 }')"
  if [ "${SCORE_INT}" -lt "${SOFT_TARGET_PCT}" ]; then
    echo "${PROG}: score ${SCORE_PCT}% < soft target ${SOFT_TARGET_PCT}% (FORMULON_MUT_STRICT was set)"
    exit 1
  fi
fi

echo "${PROG}: mutation report rendered (logs: ${LOG_DIR})"
exit 0
