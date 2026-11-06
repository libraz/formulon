#!/bin/sh
# check_install.sh
#
# Local check that the shared C ABI is usable from outside the source tree:
# builds libformulon in a separate build directory, installs it to a
# temporary prefix, then builds tests/install/consumer against that prefix
# through both find_package(Formulon CONFIG) and pkg-config, and runs both
# binaries. Exits non-zero on any failure. Not run by CI.

set -eu

ROOT=$(cd "$(dirname "$0")/../.." && pwd)
BUILD_DIR="$ROOT/build-install"
WORK=$(mktemp -d "${TMPDIR:-/tmp}/formulon-install.XXXXXX")
trap 'rm -rf "$WORK"' EXIT
PREFIX="$WORK/prefix"
CONSUMER_BUILD="$WORK/consumer"

JOBS=3
command -v pkg-config >/dev/null 2>&1 || { echo "check_install: pkg-config not found" >&2; exit 1; }

LAUNCHER=""
if command -v ccache >/dev/null 2>&1; then
  LAUNCHER="-DCMAKE_C_COMPILER_LAUNCHER=ccache -DCMAKE_CXX_COMPILER_LAUNCHER=ccache"
else
  echo "check_install: ccache not found, building without it" >&2
fi

# shellcheck disable=SC2086
nice -n 10 cmake -S "$ROOT" -B "$BUILD_DIR" \
  -DCMAKE_BUILD_TYPE=Release \
  -DFM_BUILD_C_API_SHARED=ON \
  -DFM_BUILD_TESTING=OFF \
  -DFM_BUILD_CLI=OFF \
  $LAUNCHER
nice -n 10 cmake --build "$BUILD_DIR" --target formulon -j "$JOBS"
cmake --install "$BUILD_DIR" --prefix "$PREFIX"

# Only the shared library, its headers and the package files may be installed.
for f in include/formulon/formulon_c.h include/formulon/formulon_status.h \
         lib/cmake/Formulon/FormulonConfig.cmake \
         lib/cmake/Formulon/FormulonConfigVersion.cmake \
         lib/cmake/Formulon/FormulonTargets.cmake \
         lib/pkgconfig/formulon.pc; do
  [ -f "$PREFIX/$f" ] || { echo "check_install: missing $f in install prefix" >&2; exit 1; }
done
STRAY=$(cd "$PREFIX" && find . -type f \
  ! -path './include/formulon/*' ! -path './lib/cmake/Formulon/*' \
  ! -path './lib/pkgconfig/formulon.pc' ! -name 'libformulon*' | sort)
if [ -n "$STRAY" ]; then
  echo "check_install: unexpected installed files:" >&2
  echo "$STRAY" >&2
  exit 1
fi

PKG_CONFIG_PATH="$PREFIX/lib/pkgconfig" \
  nice -n 10 cmake -S "$ROOT/tests/install/consumer" -B "$CONSUMER_BUILD" \
  -DCMAKE_BUILD_TYPE=Release \
  -DCMAKE_PREFIX_PATH="$PREFIX"
nice -n 10 cmake --build "$CONSUMER_BUILD" -j "$JOBS"

"$CONSUMER_BUILD/consumer_cmake"
"$CONSUMER_BUILD/consumer_pkgconfig"
echo "check_install: ok"
