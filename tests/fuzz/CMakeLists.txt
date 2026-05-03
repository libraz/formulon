# tests/fuzz/CMakeLists.txt
#
# libFuzzer harnesses for the parser, evaluator, and OOXML reader.
#
# Gated on FM_BUILD_FUZZ=ON and a Clang-family compiler. On any other
# toolchain we emit a STATUS message and return early so configuring with
# FM_BUILD_FUZZ=ON on GCC / MSVC does not fail the build.
#
# Apple clang's libFuzzer support varies between Xcode versions; if the
# host's clang lacks libFuzzer the configure still succeeds but the
# generated link step will fail at build time. That is acceptable: the
# parent never builds these targets in default CI, only in the dedicated
# nightly fuzz runner.

if(NOT FM_BUILD_FUZZ)
  return()
endif()

if(NOT CMAKE_CXX_COMPILER_ID MATCHES "Clang")
  message(STATUS "Formulon fuzz: skipped (compiler '${CMAKE_CXX_COMPILER_ID}' is not Clang)")
  return()
endif()

message(STATUS "Formulon fuzz: enabling libFuzzer harnesses")

set(_FORMULON_FUZZ_FLAGS
  -fsanitize=fuzzer,address,undefined
  -fno-exceptions
  -fno-rtti
)

# Per-harness wiring. Each harness:
#   * links to formulon_static for engine code
#   * compiles + links with -fsanitize=fuzzer,address,undefined
#   * lands in ${CMAKE_BINARY_DIR}/bin alongside the CLI
#   * registers a smoke ctest with -runs=1000 -timeout=10 and label SLOW
function(_formulon_add_fuzz_target name source)
  add_executable(${name} ${source})
  target_link_libraries(${name} PRIVATE formulon_static)
  target_include_directories(${name} PRIVATE ${CMAKE_SOURCE_DIR}/src)
  target_compile_options(${name} PRIVATE ${_FORMULON_FUZZ_FLAGS})
  target_link_options(${name} PRIVATE ${_FORMULON_FUZZ_FLAGS})
  set_target_properties(${name} PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin"
  )
endfunction()

_formulon_add_fuzz_target(parser_fuzz parser_fuzz.cpp)
_formulon_add_fuzz_target(eval_fuzz   eval_fuzz.cpp)
_formulon_add_fuzz_target(xlsx_fuzz   xlsx_fuzz.cpp)

# Smoke tests run a handful of iterations against an empty corpus so a
# default `ctest` invocation surfaces obvious regressions without
# committing to a full nightly fuzz session. The 6h nightly runs are
# scheduled out-of-band by the parent infrastructure with
# `-max_total_time=21600`.
add_test(
  NAME FuzzParserSmoke
  COMMAND parser_fuzz -runs=1000 -timeout=10 -dict=${CMAKE_CURRENT_SOURCE_DIR}/parser_dict.txt
)
set_tests_properties(FuzzParserSmoke PROPERTIES LABELS "SLOW")

add_test(
  NAME FuzzEvalSmoke
  COMMAND eval_fuzz -runs=1000 -timeout=10 -dict=${CMAKE_CURRENT_SOURCE_DIR}/parser_dict.txt
)
set_tests_properties(FuzzEvalSmoke PROPERTIES LABELS "SLOW")

add_test(
  NAME FuzzXlsxSmoke
  COMMAND xlsx_fuzz -runs=1000 -timeout=10
)
set_tests_properties(FuzzXlsxSmoke PROPERTIES LABELS "SLOW")
