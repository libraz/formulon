#!/bin/sh
# run_tsan.sh
#
# ThreadSanitizer CI gate. Builds the concurrency-test executables under
# `-fsanitize=thread`, runs them via ctest filtered to the `TSAN` label,
# and exits non-zero if any `WARNING: ThreadSanitizer` or `data race` line
# appears in the captured output.
#
# This script is intentionally orthogonal to the default `make build` /
# `make test` flow: TSan instrumentation is incompatible with the
# coverage / asan / ubsan builds and runs an order of magnitude slower
# than the standard test suite. It belongs in a dedicated CI job.
#
# Usage:
#   tools/ci/run_tsan.sh
#
# Environment overrides:
#   FORMULON_TSAN_BUILD_DIR  Build directory (default: build-tsan)
#   FORMULON_TSAN_TIMEOUT    Per-test timeout in seconds (default: 300)
#   FORMULON_TSAN_LABEL      ctest label to select (default: TSAN)
#   CMAKE                    cmake binary (default: cmake)
#   CTEST                    ctest binary (default: ctest)
#   CMAKE_BUILD_TYPE         Build type (default: RelWithDebInfo)
#   CXX_COMPILER             Override CXX (default: leave to CMake's auto-detect)
#
# Exit codes:
#   0 -- every TSan test passed AND no race warnings were emitted
#   1 -- one or more tests failed OR TSan emitted a race warning
#   2 -- build / configure error
#
# TODO: GitHub Actions wiring (a dedicated `tsan` job in
# `.github/workflows/ci.yml`) is intentionally deferred to a separate
# bundle so the wiring decision (job dependencies, runner os, caching
# strategy) can be evaluated against the WASM and native jobs together.
# Until then, this script is the canonical way to drive the gate locally
# or from a custom CI runner.

set -eu

PROG="$(basename "$0")"

BUILD_DIR="${FORMULON_TSAN_BUILD_DIR:-build-tsan}"
TIMEOUT_SEC="${FORMULON_TSAN_TIMEOUT:-300}"
LABEL="${FORMULON_TSAN_LABEL:-TSAN}"
CMAKE="${CMAKE:-cmake}"
CTEST="${CTEST:-ctest}"
BUILD_TYPE="${CMAKE_BUILD_TYPE:-RelWithDebInfo}"

# Find the repo root: assume this script lives at <repo>/tools/ci/run_tsan.sh.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../.." && pwd)"
cd "${REPO_ROOT}"

LOG_DIR="${REPO_ROOT}/${BUILD_DIR}/_tsan_logs"
mkdir -p "${LOG_DIR}"
CONFIG_LOG="${LOG_DIR}/configure.log"
BUILD_LOG="${LOG_DIR}/build.log"
RUN_LOG="${LOG_DIR}/ctest.log"

# Determine compiler (Apple Clang ships TSan on macOS; on Linux we prefer
# clang if available, otherwise fall back to whatever CMake picks).
if [ -z "${CXX_COMPILER:-}" ]; then
  if [ "$(uname -s)" = "Darwin" ]; then
    # Apple Clang has full TSan support; rely on the default toolchain.
    CXX_COMPILER=""
  elif command -v clang++ >/dev/null 2>&1; then
    CXX_COMPILER="clang++"
  else
    CXX_COMPILER=""  # Let CMake decide.
  fi
fi

# TSan-specific flags. `-O1 -g` is the canonical recipe: optimisation
# light enough to keep the runtime overhead bounded, debug info dense
# enough that race reports point at real source lines.
SAN_CXX_FLAGS="-fsanitize=thread -O1 -g -fno-omit-frame-pointer"
SAN_LD_FLAGS="-fsanitize=thread"

echo "${PROG}: TSan build dir: ${BUILD_DIR}"
echo "${PROG}: configuring..."

# Pass -DCMAKE_*_FLAGS as single arguments so the embedded spaces in
# `-O1 -g -fno-omit-frame-pointer` survive shell word-splitting. Using
# positional parameters avoids the quoting traps of a single string var.
set -- \
  -S "${REPO_ROOT}" \
  -B "${BUILD_DIR}" \
  -DCMAKE_BUILD_TYPE="${BUILD_TYPE}" \
  -DCMAKE_CXX_FLAGS="${SAN_CXX_FLAGS}" \
  -DCMAKE_EXE_LINKER_FLAGS="${SAN_LD_FLAGS}" \
  -DCMAKE_SHARED_LINKER_FLAGS="${SAN_LD_FLAGS}" \
  -DFM_BUILD_CLI=OFF
if [ -n "${CXX_COMPILER}" ]; then
  set -- "$@" -DCMAKE_CXX_COMPILER="${CXX_COMPILER}"
fi

if ! "${CMAKE}" "$@" > "${CONFIG_LOG}" 2>&1; then
  echo "${PROG}: cmake configure failed; see ${CONFIG_LOG}" >&2
  tail -50 "${CONFIG_LOG}" >&2 || true
  exit 2
fi

echo "${PROG}: building concurrency tests (target: formulon_concurrency_tsan_tests)..."
if ! ${CMAKE} --build "${BUILD_DIR}" --target formulon_concurrency_tsan_tests --parallel \
    > "${BUILD_LOG}" 2>&1; then
  echo "${PROG}: build failed; see ${BUILD_LOG}" >&2
  tail -100 "${BUILD_LOG}" >&2 || true
  exit 2
fi

# TSan options: halt on first error so a race report points at the
# triggering test rather than the cumulative noise. `report_thread_leaks=0`
# matches Google's default for libraries that intentionally hold worker
# threads across tests; we still join every thread we spawn, but it
# silences harmless atexit-during-pool noise on some libstdc++ versions.
TSAN_OPTIONS_DEFAULT="halt_on_error=1:second_deadlock_stack=1:report_thread_leaks=0"
TSAN_OPTIONS="${TSAN_OPTIONS:-${TSAN_OPTIONS_DEFAULT}}"
export TSAN_OPTIONS

echo "${PROG}: running ctest -L ${LABEL} (timeout ${TIMEOUT_SEC}s per test)..."
set +e
(cd "${BUILD_DIR}" && ${CTEST} -L "${LABEL}" --output-on-failure --timeout "${TIMEOUT_SEC}") \
    > "${RUN_LOG}" 2>&1
CTEST_RC=$?
set -e

# Inspect the captured output for TSan race warnings. We grep AFTER the
# run completes (the long-running-command policy in CLAUDE.md forbids
# piping live output through grep).
RACE_HITS=$(grep -E "WARNING: ThreadSanitizer|data race" "${RUN_LOG}" | wc -l | tr -d ' ')

if [ "${CTEST_RC}" -ne 0 ]; then
  echo "${PROG}: ctest exited non-zero (rc=${CTEST_RC}); see ${RUN_LOG}" >&2
  tail -200 "${RUN_LOG}" >&2 || true
fi

if [ "${RACE_HITS}" -ne 0 ]; then
  echo "${PROG}: TSan reported ${RACE_HITS} race line(s); see ${RUN_LOG}" >&2
  grep -E "WARNING: ThreadSanitizer|data race|^  #" "${RUN_LOG}" >&2 || true
  exit 1
fi

if [ "${CTEST_RC}" -ne 0 ]; then
  exit 1
fi

echo "${PROG}: TSan gate passed (label=${LABEL}, races=0)"
exit 0
