# FetchPugixml.cmake
#
# Downloads and configures pugixml v1.14 via FetchContent.
# pugixml parses OOXML parts (sheet1.xml, workbook.xml, styles.xml, ...).
# We enable PUGIXML_COMPACT for WASM size and PUGIXML_NO_EXCEPTIONS to
# match the project-wide -fno-exceptions policy.

include(FetchContent)

# Upstream reads these as CMake options in its own CMakeLists and folds them
# into PUGIXML_PUBLIC_DEFINITIONS on the ``pugixml-static`` target, so they
# propagate to consumers via the ``pugixml::pugixml`` alias.
set(PUGIXML_COMPACT       ON CACHE BOOL "" FORCE)
set(PUGIXML_NO_EXCEPTIONS ON CACHE BOOL "" FORCE)
# Ensure only the static library is built (saves time and prevents a
# superfluous SHARED target from appearing in the Formulon build graph).
set(BUILD_SHARED_LIBS OFF)
set(PUGIXML_BUILD_SHARED_AND_STATIC_LIBS OFF CACHE BOOL "" FORCE)

# pugixml 1.14 still declares cmake_minimum_required(VERSION 3.5); CMake 4.x
# removed that compatibility.  FetchMiniz.cmake already sets this, but guard
# against include order differences.
if(NOT DEFINED CMAKE_POLICY_VERSION_MINIMUM)
  set(CMAKE_POLICY_VERSION_MINIMUM 3.5)
endif()

FetchContent_Declare(
  pugixml
  GIT_REPOSITORY https://github.com/zeux/pugixml.git
  GIT_TAG v1.14
  GIT_SHALLOW TRUE
)

FetchContent_MakeAvailable(pugixml)

# Upstream target layout (v1.14):
#   pugixml-static       — STATIC library (the one actually built here).
#   pugixml              — INTERFACE library that links to pugixml-static.
#   pugixml::pugixml     — ALIAS to the INTERFACE target (the recommended
#                          consumer-facing target).
# ALIAS-of-ALIAS is disallowed, so we alias directly to the underlying
# non-ALIAS target.
if(NOT TARGET formulon::pugixml)
  if(TARGET pugixml-static)
    add_library(formulon::pugixml ALIAS pugixml-static)
  elseif(TARGET pugixml)
    add_library(formulon::pugixml ALIAS pugixml)
  else()
    message(FATAL_ERROR "FetchPugixml: upstream did not expose a recognised target")
  endif()
endif()
