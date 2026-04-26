# tests/oracle/CMakeLists.txt
#
# Defines the formulon_oracle_tests executable. The target is parameterized
# over the golden JSON files under tests/oracle/golden/ which are produced
# by tools/oracle/oracle_gen.py on macOS + Excel 365. When no golden files
# exist (the default state before oracle-gen has been run), the test binary
# links and runs but registers zero parameterized cases — ctest output is
# clean and the build stays green.

set(FORMULON_ORACLE_TEST_SOURCES
  json_reader.cpp
  oracle_runner.cpp
  oracle_test.cpp
)

add_executable(formulon_oracle_tests ${FORMULON_ORACLE_TEST_SOURCES})

target_link_libraries(formulon_oracle_tests
  PRIVATE
    formulon_static
    gtest_main
)

target_include_directories(formulon_oracle_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
    ${CMAKE_SOURCE_DIR}        # for tests/oracle/*.h via tests/oracle/...
)

# Hand the golden directory to the test binary as a compile-time define so
# the parameter provider can enumerate files at process start. An absent
# directory is tolerated (load_oracle_cases returns an empty vector).
target_compile_definitions(formulon_oracle_tests
  PRIVATE
    FORMULON_ORACLE_GOLDEN_DIR="${CMAKE_CURRENT_SOURCE_DIR}/golden"
)

formulon_apply_warnings(formulon_oracle_tests)

# Register with CTest under the "oracle" label so `make oracle-verify`
# (which shells to `ctest -L oracle`) can select this binary's cases
# without also running the full unit suite.
#
# DISCOVERY_MODE PRE_TEST rediscovers parameterized cases at every `ctest`
# invocation. That matters here because golden JSON files are dynamic
# test inputs — when `make oracle-gen` writes a new suite, the test list
# must update without requiring a relink.
gtest_discover_tests(formulon_oracle_tests
  DISCOVERY_MODE PRE_TEST
  PROPERTIES LABELS "oracle"
)

# ---------------------------------------------------------------------------
# IronCalc oracle target
#
# Secondary oracle that reads golden JSON imported from the IronCalc xlsx
# fixtures (`tests/oracle/external/ironcalc/fixtures/`) via
# `tools/oracle/ironcalc_import.py`. The binary reuses oracle_test.cpp
# verbatim; only FORMULON_ORACLE_GOLDEN_DIR differs so the parameter
# provider pulls from `golden/ironcalc/` instead of `golden/`.
#
# Registered under the `ironcalc` CTest label so the primary `oracle`
# suite can still be invoked standalone by `make oracle-verify`.
# ---------------------------------------------------------------------------

add_executable(formulon_ironcalc_oracle_tests ${FORMULON_ORACLE_TEST_SOURCES})

target_link_libraries(formulon_ironcalc_oracle_tests
  PRIVATE
    formulon_static
    gtest_main
)

target_include_directories(formulon_ironcalc_oracle_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
    ${CMAKE_SOURCE_DIR}
)

target_compile_definitions(formulon_ironcalc_oracle_tests
  PRIVATE
    FORMULON_ORACLE_GOLDEN_DIR="${CMAKE_CURRENT_SOURCE_DIR}/golden/ironcalc"
)

formulon_apply_warnings(formulon_ironcalc_oracle_tests)

# Single-shot ctest entry: the IronCalc suite registers ~12,500 parameterized
# gtest cases, and registering each as its own ctest test makes `ctest` spend
# ~0.2 s of process-launch overhead per case (377 s wall total, vs ~0.2 s
# when the binary runs once and lets gtest iterate internally). For iteration
# speed we register one ctest entry that invokes the binary directly; gtest
# still prints per-case failures in `--output-on-failure`.
add_test(NAME formulon_ironcalc_oracle
         COMMAND formulon_ironcalc_oracle_tests --gtest_brief=1)
set_tests_properties(formulon_ironcalc_oracle PROPERTIES LABELS "ironcalc")

# ---------------------------------------------------------------------------
# Variant oracle target (opt-in)
#
# Loads goldens from tests/oracle/variants/<tag>/golden/ and tags each TEST_P
# parameter with `__<tag>` so primary and variant test names never collide.
# Built only when FORMULON_ORACLE_VARIANTS=ON; defaults to OFF so the
# baseline build tree stays unchanged.
#
# Variant directories are scanned at configure time. The discovered list is
# baked into the binary as a compile-time default; the env var
# FORMULON_ORACLE_VARIANT_DIRS overrides this at runtime if set.
# ---------------------------------------------------------------------------

option(FORMULON_ORACLE_VARIANTS "Build the variant oracle test binary" OFF)

if(FORMULON_ORACLE_VARIANTS)
  set(_variant_root "${CMAKE_CURRENT_SOURCE_DIR}/variants")
  set(_variant_dirs_default "")
  if(EXISTS "${_variant_root}")
    file(GLOB _variant_subdirs RELATIVE "${_variant_root}" "${_variant_root}/*")
    foreach(_tag IN LISTS _variant_subdirs)
      set(_golden "${_variant_root}/${_tag}/golden")
      if(IS_DIRECTORY "${_golden}")
        if(_variant_dirs_default STREQUAL "")
          set(_variant_dirs_default "${_tag}=${_golden}")
        else()
          set(_variant_dirs_default "${_variant_dirs_default}:${_tag}=${_golden}")
        endif()
      endif()
    endforeach()
  endif()

  add_executable(formulon_oracle_variant_tests ${FORMULON_ORACLE_TEST_SOURCES})

  target_link_libraries(formulon_oracle_variant_tests
    PRIVATE
      formulon_static
      gtest_main
  )

  target_include_directories(formulon_oracle_variant_tests
    PRIVATE
      ${CMAKE_SOURCE_DIR}/src
      ${CMAKE_SOURCE_DIR}
  )

  # Empty primary dir so the binary loads ONLY variants. variant_dirs_default
  # gets the configure-time scan; runtime env var still overrides.
  target_compile_definitions(formulon_oracle_variant_tests
    PRIVATE
      FORMULON_ORACLE_GOLDEN_DIR=""
      FORMULON_ORACLE_VARIANT_DIRS_DEFAULT="${_variant_dirs_default}"
  )

  formulon_apply_warnings(formulon_oracle_variant_tests)

  gtest_discover_tests(formulon_oracle_variant_tests
    DISCOVERY_MODE PRE_TEST
    PROPERTIES LABELS "VARIANT"
  )
endif()
