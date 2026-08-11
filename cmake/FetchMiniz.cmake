# FetchMiniz.cmake
#
# Downloads and configures miniz v3.1.2 via FetchContent.
# miniz provides zip/deflate support for OOXML package (M6) and XLSB
# compression. Upstream: https://github.com/richgel999/miniz.
#
# Upstream ships its own CMakeLists which builds a static library. We turn
# off its tests, examples, and install target to keep the build lean.

include(FetchContent)

# miniz 3.1.2 ships a CMakeLists that still declares compatibility with
# CMake < 3.5, which CMake 4.x removed. Opt-in to the legacy policy version
# just for the fetched subtree.
if(NOT DEFINED CMAKE_POLICY_VERSION_MINIMUM)
  set(CMAKE_POLICY_VERSION_MINIMUM 3.5)
endif()

# miniz 3.1.2 uses these option names directly (not MINIZ_* aliases).
set(BUILD_TESTS    OFF CACHE BOOL "" FORCE)
set(BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
set(BUILD_FUZZERS  OFF CACHE BOOL "" FORCE)
set(INSTALL_PROJECT OFF CACHE BOOL "" FORCE)
set(BUILD_HEADER_ONLY OFF CACHE BOOL "" FORCE)
set(AMALGAMATE_SOURCES OFF CACHE BOOL "" FORCE)

FetchContent_Declare(
  miniz
  GIT_REPOSITORY https://github.com/richgel999/miniz.git
  GIT_TAG 3.1.2
  GIT_SHALLOW TRUE
)

FetchContent_MakeAvailable(miniz)

# Upstream target name is ``miniz`` (static library) as of 3.1.2. If a newer
# tag renames it, adjust here.
if(TARGET miniz)
  set(_FORMULON_MINIZ_TARGET miniz)
elseif(TARGET minizc)
  set(_FORMULON_MINIZ_TARGET minizc)
else()
  message(FATAL_ERROR "FetchMiniz: upstream did not expose a recognised target")
endif()

if(NOT TARGET formulon::miniz)
  add_library(formulon::miniz ALIAS ${_FORMULON_MINIZ_TARGET})
endif()

# miniz's public headers contain optional static zlib-compatibility wrappers.
# They remain available to callers (the smoke tests use them), but Clang's
# -Wunused-function reports the unused wrappers when a strict consumer merely
# includes `miniz.h`. Classify only this fetched dependency's include paths as
# SYSTEM so the project-wide warning policy remains strict for Formulon's own
# headers and sources.
get_target_property(_FORMULON_MINIZ_INCLUDE_DIRS ${_FORMULON_MINIZ_TARGET} INTERFACE_INCLUDE_DIRECTORIES)
if(_FORMULON_MINIZ_INCLUDE_DIRS AND NOT _FORMULON_MINIZ_INCLUDE_DIRS STREQUAL "_FORMULON_MINIZ_INCLUDE_DIRS-NOTFOUND")
  target_include_directories(${_FORMULON_MINIZ_TARGET} SYSTEM INTERFACE
                             ${_FORMULON_MINIZ_INCLUDE_DIRS})
else()
  # The 3.1.x CMake target can omit its header directory when consumed
  # through FetchContent. Keep both the source and generated-header paths
  # available to object-library consumers (notably the sanitizer-only fuzz
  # core) as system paths.
  target_include_directories(${_FORMULON_MINIZ_TARGET} SYSTEM INTERFACE
                             "${miniz_SOURCE_DIR}" "${miniz_BINARY_DIR}")
endif()
