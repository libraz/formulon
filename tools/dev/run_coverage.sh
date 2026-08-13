#!/bin/sh
# run_coverage.sh
#
# Local coverage diagnostic. Builds with the gcov-instrumented
# `coverage` CMake preset, runs the fast ctest labels, captures an lcov
# tracefile, filters out third-party / generated / test sources, and
# delegates to tools/dev/coverage_report.py for slicing + reporting.
#
# This is a *local* diagnostic, not a CI gate. CI is verification, not
# the development loop, and coverage runs are too slow (multi-minute
# instrumented build + run) to gate either on. Run this script when
# investigating gaps -- typically before a milestone, when adding a
# new module, or when triaging a "what's untested?" question.
#
# Per-area observed targets (defined in CLAUDE.md, aspirational only):
#   util/value/eval (excluding function families)  ~95%
#   functions       (builtins / lookups / criteria) ~98%
#   io              (OOXML / XLSB / CSV)            ~90%
#
# By default the script prints the per-area table and the list of
# zero-coverage files (the genuinely useful signal -- files no test
# touches). Set FORMULON_COV_STRICT=1 to opt into a non-zero exit when
# any gated area falls below its target; this is for ad-hoc local
# enforcement and is never used by CI.
#
# Usage:
#   tools/dev/run_coverage.sh
#   FORMULON_COV_STRICT=1 tools/dev/run_coverage.sh   # opt-in failure
#
# Environment overrides:
#   FORMULON_COV_BUILD_DIR  Build directory (default: build-coverage)
#   FORMULON_COV_TIMEOUT    Per-test timeout in seconds (default: 60)
#   FORMULON_COV_STRICT     If non-empty, exit 1 when any gated area
#                           falls below its target (default: report only)
#   LCOV                    lcov binary (default: lcov)
#   CMAKE                   cmake binary (default: cmake)
#   CTEST                   ctest binary (default: ctest)
#
# Exit codes:
#   0 -- coverage captured and report rendered (regardless of targets,
#        unless FORMULON_COV_STRICT is set)
#   1 -- FORMULON_COV_STRICT was set and at least one gated area was
#        below its target
#   2 -- configure / build / tooling-missing failure

set -eu

PROG="$(basename "$0")"

BUILD_DIR="${FORMULON_COV_BUILD_DIR:-build-coverage}"
TIMEOUT_SEC="${FORMULON_COV_TIMEOUT:-60}"
STRICT_FLAG="${FORMULON_COV_STRICT:-}"
LCOV="${LCOV:-lcov}"
CMAKE="${CMAKE:-cmake}"
CTEST="${CTEST:-ctest}"

# Find the repo root: assume this script lives at <repo>/tools/dev/run_coverage.sh.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../.." && pwd)"
cd "${REPO_ROOT}"

LOG_DIR="${REPO_ROOT}/${BUILD_DIR}/_cov_logs"
mkdir -p "${LOG_DIR}"
CONFIG_LOG="${LOG_DIR}/configure.log"
BUILD_LOG="${LOG_DIR}/build.log"
RUN_LOG="${LOG_DIR}/ctest.log"
LCOV_LOG="${LOG_DIR}/lcov.log"
REPORT_LOG="${LOG_DIR}/report.log"

TRACEFILE_RAW="${LOG_DIR}/coverage.info"
TRACEFILE_FILTERED="${LOG_DIR}/coverage.filtered.info"

# Tooling preflight. lcov drives gcov internally, so a missing gcov
# surfaces here too via lcov's own probe; we still check both for a
# friendlier diagnostic.
if ! command -v "${LCOV}" >/dev/null 2>&1; then
  echo "${PROG}: '${LCOV}' not found in PATH; install lcov (apt: lcov, brew: lcov)" >&2
  exit 2
fi
if ! command -v gcov >/dev/null 2>&1; then
  echo "${PROG}: 'gcov' not found in PATH; install a gcc/g++ toolchain" >&2
  exit 2
fi

echo "${PROG}: coverage build dir: ${BUILD_DIR}"
echo "${PROG}: configuring (preset=coverage)..."

# The `coverage` configure preset is declared in CMakePresets.json and
# wires up `--coverage -fprofile-arcs -ftest-coverage` plus the matching
# linker flags.
if ! "${CMAKE}" --preset coverage > "${CONFIG_LOG}" 2>&1; then
  echo "${PROG}: cmake configure failed; see ${CONFIG_LOG}" >&2
  tail -50 "${CONFIG_LOG}" >&2 || true
  exit 2
fi

# Build everything. tests/CMakeLists.txt reaches several executables
# (unit, integration, c_api, cli, oracle, ironcalc-oracle, parity,
# cf-oracle, ...). Skipping any of them under-counts coverage for
# whatever sources their test driver alone exercises.
echo "${PROG}: building all test executables..."
if ! "${CMAKE}" --build "${BUILD_DIR}" --parallel > "${BUILD_LOG}" 2>&1; then
  echo "${PROG}: build failed; see ${BUILD_LOG}" >&2
  tail -100 "${BUILD_LOG}" >&2 || true
  exit 2
fi

# Reset gcov counters so leftover data from a cached build directory
# can't inflate the captured tracefile.
echo "${PROG}: zeroing counters..."
if ! "${LCOV}" --zerocounters --directory "${BUILD_DIR}" >> "${LCOV_LOG}" 2>&1; then
  echo "${PROG}: lcov --zerocounters failed; see ${LCOV_LOG}" >&2
  tail -50 "${LCOV_LOG}" >&2 || true
  exit 2
fi

# Run the fast test labels. Coverage data is still useful even when a
# test fails, so we capture rc and continue rather than early-exit;
# the native CI job already enforces test correctness.
echo "${PROG}: running ctest -LE 'SLOW|BENCH|TSAN' (timeout ${TIMEOUT_SEC}s per test)..."
set +e
(cd "${BUILD_DIR}" && "${CTEST}" -LE "SLOW|BENCH|TSAN" --output-on-failure --timeout "${TIMEOUT_SEC}") \
    > "${RUN_LOG}" 2>&1
CTEST_RC=$?
set -e

if [ "${CTEST_RC}" -ne 0 ]; then
  echo "${PROG}: ctest exited rc=${CTEST_RC} (continuing for coverage capture); see ${RUN_LOG}" >&2
  tail -100 "${RUN_LOG}" >&2 || true
fi

echo "${PROG}: capturing lcov tracefile..."
# `--ignore-errors mismatch,negative,unused,inconsistent` accommodates
# common lcov diagnostics on modern toolchains:
#   mismatch     - .gcno / .gcda function-id drift across compilers
#   negative     - rare counter underflow on heavily-templated TUs
#   unused       - lcov complaining about a known-empty filter list
#   inconsistent - branch-vs-line counter disagreement (we disable
#                  branch coverage anyway via lcov_branch_coverage=0)
if ! "${LCOV}" --capture \
      --directory "${BUILD_DIR}" \
      --output-file "${TRACEFILE_RAW}" \
      --rc lcov_branch_coverage=0 \
      --ignore-errors mismatch,negative,unused,inconsistent \
      >> "${LCOV_LOG}" 2>&1; then
  echo "${PROG}: lcov --capture failed; see ${LCOV_LOG}" >&2
  tail -100 "${LCOV_LOG}" >&2 || true
  exit 2
fi

echo "${PROG}: filtering tracefile..."
if ! "${LCOV}" --remove "${TRACEFILE_RAW}" \
      '/usr/*' \
      '*/third_party/*' \
      '*/_deps/*' \
      '*/tests/*' \
      '*/build*/*' \
      --output-file "${TRACEFILE_FILTERED}" \
      --rc lcov_branch_coverage=0 \
      --ignore-errors mismatch,negative,unused,inconsistent \
      >> "${LCOV_LOG}" 2>&1; then
  echo "${PROG}: lcov --remove failed; see ${LCOV_LOG}" >&2
  tail -100 "${LCOV_LOG}" >&2 || true
  exit 2
fi

echo "${PROG}: rendering coverage report..."

set -- "${TRACEFILE_FILTERED}"
if [ -n "${STRICT_FLAG}" ]; then
  set -- "$@" --strict
fi

set +e
python3 "${REPO_ROOT}/tools/dev/coverage_report.py" "$@" > "${REPORT_LOG}" 2>&1
REPORT_RC=$?
set -e

# Always echo the report so the user sees it.
cat "${REPORT_LOG}"

# Also emit a json copy for ad-hoc tooling that wants per-area numbers.
# Best-effort; does not affect rc.
python3 "${REPO_ROOT}/tools/dev/coverage_report.py" \
    "${TRACEFILE_FILTERED}" --json \
    > "${LOG_DIR}/coverage.json" 2>/dev/null || true

# coverage_report.py exits 2 for tracefile/argument errors, 1 for
# below-target with --strict, 0 otherwise.
if [ "${REPORT_RC}" -eq 2 ]; then
  echo "${PROG}: coverage_report.py reported tracefile / argument error" >&2
  exit 2
fi

if [ "${REPORT_RC}" -eq 1 ]; then
  # Only reachable when FORMULON_COV_STRICT is set.
  echo "${PROG}: at least one area below target (FORMULON_COV_STRICT was set)"
  exit 1
fi

echo "${PROG}: coverage report rendered (logs: ${LOG_DIR})"
exit 0
