# FetchMiniz.cmake
#
# Downloads and configures miniz v3.0.2 via FetchContent.
# miniz provides zip/deflate support for OOXML package (M6) and XLSB
# compression. Upstream: https://github.com/richgel999/miniz.
#
# Upstream ships its own CMakeLists which builds a static library. We turn
# off its tests, examples, and install target to keep the build lean.

include(FetchContent)

# miniz 3.0.2 ships a CMakeLists that still declares compatibility with
# CMake < 3.5, which CMake 4.x removed. Opt-in to the legacy policy version
# just for the fetched subtree.
if(NOT DEFINED CMAKE_POLICY_VERSION_MINIMUM)
  set(CMAKE_POLICY_VERSION_MINIMUM 3.5)
endif()

set(MINIZ_BUILD_TESTS    OFF CACHE BOOL "" FORCE)
set(MINIZ_BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
set(MINIZ_INSTALL_PROJECT OFF CACHE BOOL "" FORCE)
# Defensive: some miniz versions use alternate names for these options.
set(BUILD_TESTS    OFF CACHE BOOL "" FORCE)
set(BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
set(BUILD_FUZZERS  OFF CACHE BOOL "" FORCE)
set(BUILD_HEADER   OFF CACHE BOOL "" FORCE)
set(AMALGAMATE_SOURCES OFF CACHE BOOL "" FORCE)

FetchContent_Declare(
  miniz
  GIT_REPOSITORY https://github.com/richgel999/miniz.git
  GIT_TAG 3.0.2
  GIT_SHALLOW TRUE
)

FetchContent_MakeAvailable(miniz)

# Upstream target name is ``miniz`` (static library) as of 3.0.2. If a newer
# tag renames it, adjust here.
if(TARGET miniz)
  if(NOT TARGET formulon::miniz)
    add_library(formulon::miniz ALIAS miniz)
  endif()
elseif(TARGET minizc)
  if(NOT TARGET formulon::miniz)
    add_library(formulon::miniz ALIAS minizc)
  endif()
else()
  message(FATAL_ERROR "FetchMiniz: upstream did not expose a recognised target")
endif()
