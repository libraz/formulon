# FetchNodeAddonApi.cmake
#
# Downloads the C++ wrapper for N-API (`napi.h`) via FetchContent.
# Header-only on top of the Node N-API C headers fetched by
# `FetchNodeApiHeaders.cmake`.
#
# Upstream: https://github.com/nodejs/node-addon-api
#
# Pinned to v8.5.0 (the v8 line is current LTS-track as of 2026-05).
# The macros `NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS` and
# `NAPI_DISABLE_CPP_EXCEPTIONS` are propagated by the consumer
# (`FormulonNodeAddon.cmake`) so this fetch step is just include-path
# wiring.

include(FetchContent)

if(NOT DEFINED CMAKE_POLICY_VERSION_MINIMUM)
  set(CMAKE_POLICY_VERSION_MINIMUM 3.5)
endif()

FetchContent_Declare(
  node_addon_api
  GIT_REPOSITORY https://github.com/nodejs/node-addon-api.git
  GIT_TAG v8.5.0
  GIT_SHALLOW TRUE
)

# Upstream has no top-level CMakeLists.txt; FetchContent_MakeAvailable
# still populates the source tree and just skips the add_subdirectory step.
FetchContent_MakeAvailable(node_addon_api)

if(NOT TARGET formulon::node_addon_api_headers)
  add_library(formulon_node_addon_api_headers INTERFACE)
  # Upstream lays the headers (`napi.h`, `napi-inl.h`) at the repo root.
  target_include_directories(formulon_node_addon_api_headers INTERFACE
    "${node_addon_api_SOURCE_DIR}"
  )
  add_library(formulon::node_addon_api_headers ALIAS formulon_node_addon_api_headers)
endif()
