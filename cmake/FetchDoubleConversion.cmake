# FetchDoubleConversion.cmake
#
# Downloads and configures Google's `double-conversion` v3.3.0 via
# FetchContent. The library provides Grisu3-based shortest-roundtrip
# double-to-string conversion (and the inverse string-to-double path),
# which Excel itself uses for numeric serialisation in OOXML cell
# values. Adopting it lets the writer emit numbers in the same shape as
# Excel-authored XLSX files, minimising diffs in roundtrip and visual
# review workflows.
#
# Upstream: https://github.com/google/double-conversion
#
# We pin to v3.3.0 (released 2023-11-22, the latest stable line as of
# late 2025) and link the static library only. Upstream ships a single
# CMake target named `double-conversion` that owns both the headers
# (under `<root>/double-conversion/*.h`) and the compiled archive.
# We alias it as `formulon::double_conversion`.

include(FetchContent)

# v3.3.0 declares cmake_minimum_required(VERSION 3.0); CMake 4.x removed
# that compatibility shim. Match the policy raise the other fetched
# dependencies use.
if(NOT DEFINED CMAKE_POLICY_VERSION_MINIMUM)
  set(CMAKE_POLICY_VERSION_MINIMUM 3.5)
endif()

# --- Static-only build -----------------------------------------------------
# Static linkage matches the rest of the engine and keeps the WASM
# artifact self-contained.
set(BUILD_SHARED_LIBS OFF CACHE BOOL "" FORCE)

# --- Tests / examples ------------------------------------------------------
# Upstream's `BUILD_TESTING` option turns on a test/ subdirectory that
# pulls in additional sources. We never run it from inside the Formulon
# build tree.
set(BUILD_TESTING OFF CACHE BOOL "" FORCE)

FetchContent_Declare(
  double_conversion
  GIT_REPOSITORY https://github.com/google/double-conversion.git
  GIT_TAG v3.3.0
  GIT_SHALLOW TRUE
)

FetchContent_MakeAvailable(double_conversion)

# Upstream target name is `double-conversion` (a hyphen, not an
# underscore). Alias under the `formulon::` namespace so downstream
# targets do not have to know whether we vendored or located a system
# package.
if(NOT TARGET formulon::double_conversion)
  if(TARGET double-conversion)
    add_library(formulon::double_conversion ALIAS double-conversion)
  elseif(TARGET double_conversion)
    add_library(formulon::double_conversion ALIAS double_conversion)
  else()
    message(FATAL_ERROR
      "FetchDoubleConversion: upstream did not expose a recognised target")
  endif()
endif()
