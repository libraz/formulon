# FetchNodeApiHeaders.cmake
#
# Downloads the official Node.js N-API C header set via FetchContent.
# The headers are pure C (`node_api.h`, `js_native_api.h`, ...) and
# describe the stable ABI that Formulon's N-API addon links against.
#
# Upstream: https://github.com/nodejs/node-api-headers
#
# Pinned to v1.5.0 (a tagged release) rather than tracking `main` so the
# fetched ABI is reproducible. When bumping, verify the addon still
# compiles cleanly under `NAPI_VERSION=8`.
#
# The repository is a pure header drop with no CMakeLists, so we expose
# the include directory as an INTERFACE library named
# `formulon::node_api_headers`. Consumers add `target_link_libraries(...
# PRIVATE formulon::node_api_headers)` to pick up the include path.

include(FetchContent)

if(NOT DEFINED CMAKE_POLICY_VERSION_MINIMUM)
  set(CMAKE_POLICY_VERSION_MINIMUM 3.5)
endif()

FetchContent_Declare(
  node_api_headers
  GIT_REPOSITORY https://github.com/nodejs/node-api-headers.git
  GIT_TAG v1.5.0
  GIT_SHALLOW TRUE
)

# Upstream has no CMakeLists.txt; FetchContent_MakeAvailable still
# populates the source tree and just skips the add_subdirectory step.
FetchContent_MakeAvailable(node_api_headers)

if(NOT TARGET formulon::node_api_headers)
  add_library(formulon_node_api_headers INTERFACE)
  target_include_directories(formulon_node_api_headers INTERFACE
    "${node_api_headers_SOURCE_DIR}/include"
  )
  add_library(formulon::node_api_headers ALIAS formulon_node_api_headers)
endif()
