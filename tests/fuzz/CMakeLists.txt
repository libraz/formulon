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
# generated link step will fail at build time. That is acceptable: these
# targets are opt-in and are not part of a default build.

if(NOT FM_BUILD_FUZZ)
  return()
endif()

if(NOT CMAKE_CXX_COMPILER_ID MATCHES "Clang")
  message(STATUS "Formulon fuzz: skipped (compiler '${CMAKE_CXX_COMPILER_ID}' is not Clang)")
  return()
endif()

# `MATCHES "Clang"` is a substring match, so `AppleClang` satisfies it -- and
# Apple's toolchain ships no `libclang_rt.fuzzer_osx.a`. Without this probe
# the configure succeeds, the whole core is rebuilt with sanitizers, and the
# first harness fails at link minutes later on a missing runtime. Link a
# trivial harness now so the answer costs seconds and names the cause.
include(CMakePushCheckState)
include(CheckCXXSourceCompiles)
cmake_push_check_state(RESET)
set(CMAKE_REQUIRED_FLAGS "-fsanitize=fuzzer")
check_cxx_source_compiles(
  "#include <cstddef>
   #include <cstdint>
   extern \"C\" int LLVMFuzzerTestOneInput(const uint8_t*, size_t) { return 0; }"
  FORMULON_HAVE_LIBFUZZER)
cmake_pop_check_state()

if(NOT FORMULON_HAVE_LIBFUZZER)
  message(FATAL_ERROR
    "FM_BUILD_FUZZ=ON but '${CMAKE_CXX_COMPILER}' cannot link -fsanitize=fuzzer.\n"
    "Apple's clang does not ship the libFuzzer runtime; install an LLVM that "
    "does and point CMAKE_C_COMPILER / CMAKE_CXX_COMPILER at it "
    "(for example `brew install llvm`), or configure without FM_BUILD_FUZZ.")
endif()

message(STATUS "Formulon fuzz: enabling libFuzzer harnesses")

# Sanitizer set for the harnesses. `fuzzer` is mandatory; the rest is
# overridable because a host's ASan runtime can be unusable independently of
# the code under test -- LLVM 18's ASan deadlocks inside its own shadow-memory
# init against the macOS 26 dyld, so the process never reaches main. Dropping
# to `fuzzer,undefined` keeps the harness runnable there.
set(FM_FUZZ_SANITIZERS "fuzzer,address,undefined" CACHE STRING
    "Comma-separated -fsanitize list for the fuzz harnesses")
message(STATUS "Formulon fuzz: sanitizers = ${FM_FUZZ_SANITIZERS}")

# Keep the production core completely free of fuzzer instrumentation.  An
# object library is intentionally built from the same source manifest as the
# production core, but with `fuzzer-no-link` and the selected memory/UB
# sanitizers.  Only the four harness executables below receive the libFuzzer
# runtime.  Linking a harness directly to `formulon_static` would leave the
# core objects uninstrumented (and adding sanitizer flags to
# `formulon_core` would contaminate every native/WASM/Node/Python artifact).
string(REPLACE "," ";" _FORMULON_FUZZ_SANITIZER_LIST "${FM_FUZZ_SANITIZERS}")
set(_FORMULON_FUZZ_CORE_SANITIZER_LIST fuzzer-no-link)
foreach(_sanitizer IN LISTS _FORMULON_FUZZ_SANITIZER_LIST)
  if(NOT _sanitizer STREQUAL "fuzzer" AND NOT _sanitizer STREQUAL "fuzzer-no-link" AND
     NOT _sanitizer STREQUAL "")
    list(APPEND _FORMULON_FUZZ_CORE_SANITIZER_LIST "${_sanitizer}")
  endif()
endforeach()
list(JOIN _FORMULON_FUZZ_CORE_SANITIZER_LIST "," _FORMULON_FUZZ_CORE_SANITIZERS)

set(_FORMULON_FUZZ_FLAGS
  -fsanitize=${FM_FUZZ_SANITIZERS}
  -fno-sanitize-recover=undefined
  -fno-exceptions
  -fno-rtti
)
set(_FORMULON_FUZZ_CORE_FLAGS
  -fsanitize=${_FORMULON_FUZZ_CORE_SANITIZERS}
  -fno-sanitize-recover=undefined
)

# This target is deliberately separate from `formulon_core`.  Its compile
# commands are the evidence that parser/eval/OOXML/XLSB objects are all built
# with the selected sanitizers, while the normal core compile commands retain
# their ordinary flags.
set(_FORMULON_FUZZ_CORE_SOURCES)
foreach(_core_source IN LISTS FORMULON_CORE_SOURCES)
  list(APPEND _FORMULON_FUZZ_CORE_SOURCES "${CMAKE_SOURCE_DIR}/${_core_source}")
endforeach()
add_library(formulon_fuzz_core OBJECT ${_FORMULON_FUZZ_CORE_SOURCES})
formulon_apply_common_deps(formulon_fuzz_core)
formulon_apply_warnings(formulon_fuzz_core)
target_include_directories(formulon_fuzz_core PRIVATE "${CMAKE_SOURCE_DIR}/src")
# Keep the strict project warnings on the fuzz-only core. Third-party
# dependencies are linked as separate targets and are not recompiled as part
# of this object library.
target_compile_options(formulon_fuzz_core PRIVATE ${_FORMULON_FUZZ_CORE_FLAGS})

add_library(formulon_fuzz_static STATIC $<TARGET_OBJECTS:formulon_fuzz_core>)
formulon_apply_common_deps(formulon_fuzz_static)

# Per-harness wiring. Each harness:
#   * links to the sanitizer-instrumented fuzz-only static core
#   * compiles + links with -fsanitize=fuzzer,address,undefined
#   * lands in ${CMAKE_BINARY_DIR}/bin alongside the CLI
#   * registers a smoke ctest with -runs=1000 -timeout=10 and label SLOW
function(_formulon_add_fuzz_target name source)
  add_executable(${name} ${source})
  target_link_libraries(${name} PRIVATE formulon_fuzz_static)
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
_formulon_add_fuzz_target(xlsb_fuzz   xlsb_fuzz.cpp)
_formulon_add_fuzz_target(print_settings_fuzz print_settings_fuzz.cpp)

# Smoke tests run a bounded number of iterations so a `ctest` invocation
# on a fuzz-enabled build surfaces obvious regressions in seconds. A
# longer session is the same binary with `-max_total_time=<seconds>`.
add_test(
  NAME FuzzParserSmoke
  COMMAND parser_fuzz -runs=1000 -timeout=10 -dict=${CMAKE_CURRENT_SOURCE_DIR}/parser_dict.txt
)
set_tests_properties(FuzzParserSmoke PROPERTIES LABELS "SLOW")
set_tests_properties(FuzzParserSmoke PROPERTIES ENVIRONMENT
  "UBSAN_OPTIONS=halt_on_error=1:print_stacktrace=1")

add_test(
  NAME FuzzEvalSmoke
  COMMAND eval_fuzz -runs=1000 -timeout=10 -dict=${CMAKE_CURRENT_SOURCE_DIR}/parser_dict.txt
)
set_tests_properties(FuzzEvalSmoke PROPERTIES LABELS "SLOW")
set_tests_properties(FuzzEvalSmoke PROPERTIES ENVIRONMENT
  "UBSAN_OPTIONS=halt_on_error=1:print_stacktrace=1")

# `read_ooxml` reaches no cell, shared-string or styles decoding until an
# input has cleared, in order, a ZIP end-of-central-directory record and
# central directory, `[Content_Types].xml`, a workbook content-type
# Override, `_rels/.rels` with an OfficeDocument relationship, and the
# workbook part. Random bytes do not clear the first of those, so an
# unseeded run terminates at the ZIP header on essentially every
# execution and the pipeline this harness names is never entered. Real
# packages are the only practical way in; they are staged into one
# build-tree corpus directory for the same reason the XLSB corpus below
# is (libFuzzer takes a list of inputs or a list of directories, never a
# mix).
set(_FORMULON_XLSX_CORPUS "${CMAKE_CURRENT_BINARY_DIR}/corpus_xlsx_seed")
file(MAKE_DIRECTORY "${_FORMULON_XLSX_CORPUS}")
file(GLOB _FORMULON_XLSX_SEEDS "${CMAKE_SOURCE_DIR}/tests/fixtures/excel/*.xlsx")
file(COPY ${_FORMULON_XLSX_SEEDS} DESTINATION "${_FORMULON_XLSX_CORPUS}")

add_test(
  NAME FuzzXlsxSmoke
  COMMAND xlsx_fuzz -runs=1000 -timeout=10 "${_FORMULON_XLSX_CORPUS}"
)
set_tests_properties(FuzzXlsxSmoke PROPERTIES LABELS "SLOW")
set_tests_properties(FuzzXlsxSmoke PROPERTIES ENVIRONMENT
  "UBSAN_OPTIONS=halt_on_error=1:print_stacktrace=1")

# The real XLSB fixture exercises the recorded Ptg / function coverage from
# the first run. `seeds/` carries committed starting inputs — including a
# portable deep-Ptg stream whose explicit textual prefix routes it straight to
# the Ptg decoder — as opposed to `corpus/`, which holds locally grown corpora
# and is not tracked.
#
# libFuzzer takes either a list of individual inputs to replay or a list of
# corpus directories, never a mix: naming the fixture file alongside the seeds
# directory made it exit with "Not a directory" before running anything. So
# the two are staged into one build-tree corpus directory and that single
# directory is what the harness is pointed at.
set(_FORMULON_XLSB_CORPUS "${CMAKE_CURRENT_BINARY_DIR}/corpus_xlsb_seed")
file(MAKE_DIRECTORY "${_FORMULON_XLSB_CORPUS}")
file(COPY "${CMAKE_CURRENT_SOURCE_DIR}/seeds/xlsb/" DESTINATION "${_FORMULON_XLSB_CORPUS}")
file(COPY "${CMAKE_SOURCE_DIR}/tests/fixtures/excel/xlsb_fidelity_base.xlsb"
     DESTINATION "${_FORMULON_XLSB_CORPUS}")

add_test(
  NAME FuzzXlsbSmoke
  COMMAND xlsb_fuzz -runs=1000 -timeout=10 "${_FORMULON_XLSB_CORPUS}"
)
set_tests_properties(FuzzXlsbSmoke PROPERTIES LABELS "SLOW")
set_tests_properties(FuzzXlsbSmoke PROPERTIES ENVIRONMENT
  "UBSAN_OPTIONS=halt_on_error=1:print_stacktrace=1")

# The print-settings setters are the only place caller-authored XML reaches
# pugixml through the C ABI, and each one re-parses what it stored to feed
# the paginator. The harness runs both halves.
add_test(
  NAME FuzzPrintSettingsSmoke
  COMMAND print_settings_fuzz -runs=1000 -timeout=10
)
set_tests_properties(FuzzPrintSettingsSmoke PROPERTIES LABELS "SLOW")
set_tests_properties(FuzzPrintSettingsSmoke PROPERTIES ENVIRONMENT
  "UBSAN_OPTIONS=halt_on_error=1:print_stacktrace=1")
