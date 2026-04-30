# FetchNanobench.cmake
#
# Fetches the nanobench v4.3.11 single-header microbenchmark library used
# by `tests/bench/`. Header-only, MIT-licensed; never linked into the WASM
# artifact (bench targets are guarded behind `FM_BUILD_TESTING` and only
# built when the host CMake also configures `tests/bench/CMakeLists.txt`).
#
# We do NOT call `FetchContent_MakeAvailable` because upstream's
# CMakeLists.txt builds a precompiled-implementation static library that
# would inherit `-fno-exceptions` from the parent project (nanobench's
# implementation uses `<exception>` / `<system_error>` internally and
# breaks under that flag). Instead we fetch the source tree only and
# expose the single header via an INTERFACE target. Each bench TU that
# needs the implementation defines `ANKERL_NANOBENCH_IMPLEMENT` exactly
# once and is compiled with `-fexceptions` locally.

include(FetchContent)

if(NOT DEFINED CMAKE_POLICY_VERSION_MINIMUM)
  set(CMAKE_POLICY_VERSION_MINIMUM 3.5)
endif()

FetchContent_Declare(
  nanobench
  GIT_REPOSITORY https://github.com/martinus/nanobench.git
  GIT_TAG v4.3.11
  GIT_SHALLOW TRUE
  # Source-only fetch: redirect SOURCE_SUBDIR at a path that contains no
  # CMakeLists.txt so `MakeAvailable` populates the source tree but does
  # not enter upstream's build (which would compile a precompiled
  # nanobench.cpp under the project's `-fno-exceptions` flags). The
  # bench TUs include the single header directly.
  SOURCE_SUBDIR _no_cmakelists_here
)

FetchContent_MakeAvailable(nanobench)

if(NOT TARGET formulon::nanobench)
  add_library(formulon_nanobench INTERFACE)
  target_include_directories(formulon_nanobench
    INTERFACE
      "${nanobench_SOURCE_DIR}/src/include"
  )
  add_library(formulon::nanobench ALIAS formulon_nanobench)
endif()
