# FetchPcre2.cmake
#
# Downloads and configures PCRE2 10.43 via FetchContent.
# PCRE2 backs Formulon's `REGEXTEST`, `REGEXEXTRACT`, and `REGEXREPLACE`
# Excel functions. Upstream: https://github.com/PCRE2Project/pcre2.
#
# We pin to 10.43 (released 2024-02-16, the latest stable line as of late
# 2025) and link the 8-bit static library only. JIT is intentionally
# disabled to keep the WASM artifact small and to avoid generating
# executable pages at runtime (Emscripten does not support JIT). UTF and
# UCP are enabled so Excel's `\d`, `\w`, `\s`, and Unicode property
# classes match the worksheet's UTF-8 strings.
#
# Note on PCRE2_SUPPORT_CALLOUT: callouts are a runtime feature in PCRE2
# 10.x and are not gated by a top-level CMake option (the upstream build
# always compiles callout support into pcre2_match). The wasm-size-guardian
# review listed PCRE2_SUPPORT_CALLOUT for completeness; the entry has no
# effect and is omitted here. Callouts are simply never registered by
# formulon::eval, so they cannot fire.
#
# Resource limits (PCRE2_HEAP_LIMIT, PCRE2_MATCH_LIMIT) are documented
# upstream as cache options, but in 10.43 they are *defaults* compiled
# into the library and can be overridden at runtime via
# pcre2_match_context_set_match_limit / set_depth_limit / set_heap_limit.
# The Formulon REGEX impl sets these at runtime (match_limit = 1_000_000,
# depth_limit = 10_000), so the build-time defaults are immaterial.

include(FetchContent)

# PCRE2 10.43 declares cmake_minimum_required(VERSION 3.5); CMake 4.x
# removed that compatibility shim. Match the policy raise the other
# fetched dependencies use.
if(NOT DEFINED CMAKE_POLICY_VERSION_MINIMUM)
  set(CMAKE_POLICY_VERSION_MINIMUM 3.5)
endif()

# --- Library variants ------------------------------------------------------
# Build only the 8-bit code unit width. The 16/32-bit variants would each
# add roughly the same code size again.
set(PCRE2_BUILD_PCRE2_8       ON  CACHE BOOL "" FORCE)
set(PCRE2_BUILD_PCRE2_16      OFF CACHE BOOL "" FORCE)
set(PCRE2_BUILD_PCRE2_32      OFF CACHE BOOL "" FORCE)

# --- Compiled-in features --------------------------------------------------
# JIT is off: Emscripten cannot generate executable pages, and the
# interpreter is fast enough for spreadsheet workloads. Keeping JIT off
# also strips the JIT compiler entirely, saving roughly 80-120 KB.
set(PCRE2_SUPPORT_JIT         OFF CACHE BOOL "" FORCE)

# Unicode + UCP: Excel REGEX semantics require \d / \w / \s to honor
# Unicode categories. PCRE2_SUPPORT_UNICODE turns on the UTF tables;
# the runtime path additionally sets PCRE2_UTF | PCRE2_UCP per call.
set(PCRE2_SUPPORT_UNICODE     ON  CACHE BOOL "" FORCE)

# I/O extras we never use; turning them off avoids pulling in zlib/bzip2.
set(PCRE2_SUPPORT_LIBZ        OFF CACHE BOOL "" FORCE)
set(PCRE2_SUPPORT_LIBBZ2      OFF CACHE BOOL "" FORCE)
set(PCRE2_SUPPORT_VALGRIND    OFF CACHE BOOL "" FORCE)

# --- Tools and tests we never ship ----------------------------------------
set(PCRE2_BUILD_PCRE2GREP     OFF CACHE BOOL "" FORCE)
set(PCRE2_BUILD_PCRE2TEST     OFF CACHE BOOL "" FORCE)
set(PCRE2_BUILD_TESTS         OFF CACHE BOOL "" FORCE)

# --- Static-only build ----------------------------------------------------
set(BUILD_SHARED_LIBS         OFF CACHE BOOL "" FORCE)

# Resource-limit defaults: these CMake cache vars are honored when the
# upstream CMakeLists translates them into HEAP_LIMIT / MATCH_LIMIT
# preprocessor defines. They can still be overridden at runtime through
# pcre2_match_context_set_*, which the Formulon REGEX impl does. Setting
# them here is defense-in-depth in case a future call site forgets the
# match-context override.
set(PCRE2_HEAP_LIMIT          20000000 CACHE STRING "" FORCE)
set(PCRE2_MATCH_LIMIT         10000000 CACHE STRING "" FORCE)

FetchContent_Declare(
  pcre2
  GIT_REPOSITORY https://github.com/PCRE2Project/pcre2.git
  GIT_TAG pcre2-10.43
  GIT_SHALLOW TRUE
)

FetchContent_MakeAvailable(pcre2)

# Upstream target layout (10.43): the static 8-bit library is exposed as
# `pcre2-8-static`. We alias it under the `formulon::pcre2` namespace so
# downstream targets can link without caring about the upstream name.
if(NOT TARGET formulon::pcre2)
  if(TARGET pcre2-8-static)
    add_library(formulon::pcre2 ALIAS pcre2-8-static)
  elseif(TARGET pcre2-8)
    add_library(formulon::pcre2 ALIAS pcre2-8)
  else()
    message(FATAL_ERROR "FetchPcre2: upstream did not expose a recognised 8-bit target")
  endif()
endif()
