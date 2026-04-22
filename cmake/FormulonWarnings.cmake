# FormulonWarnings.cmake
#
# Provides formulon_apply_warnings(target) which applies the strict
# project-wide warning and code-generation flags required by CLAUDE.md.
#
# - -Wall -Wextra -Wpedantic -Werror -Wshadow -Wconversion (strict warnings)
# - -fno-exceptions -fno-rtti (enforces Expected<T, Error> usage)
# - -fvisibility=hidden -fvisibility-inlines-hidden (clean WASM export surface)

function(formulon_apply_warnings target)
  if(NOT TARGET ${target})
    message(FATAL_ERROR "formulon_apply_warnings: target '${target}' does not exist")
  endif()

  get_target_property(_target_type ${target} TYPE)

  if(CMAKE_CXX_COMPILER_ID MATCHES "Clang|GNU|AppleClang")
    set(_warn_flags
      -Wall
      -Wextra
      -Wpedantic
      -Werror
      -Wshadow
      -Wconversion
    )
    set(_codegen_flags
      -fno-exceptions
      -fno-rtti
      -fvisibility=hidden
      -fvisibility-inlines-hidden
    )
  elseif(CMAKE_CXX_COMPILER_ID STREQUAL "MSVC")
    # MSVC flag mapping for potential future Windows support.
    set(_warn_flags /W4 /WX)
    set(_codegen_flags /EHs-c- /GR-)
  else()
    set(_warn_flags "")
    set(_codegen_flags "")
  endif()

  if(_target_type STREQUAL "INTERFACE_LIBRARY")
    target_compile_options(${target} INTERFACE ${_warn_flags} ${_codegen_flags})
  else()
    target_compile_options(${target} PRIVATE ${_warn_flags} ${_codegen_flags})
  endif()
endfunction()
