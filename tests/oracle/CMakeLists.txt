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
  oracle_anchor.cpp
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
# Anchor helper unit test
#
# Standalone gtest binary covering `anchor_or_self`, the spill-anchor
# projection helper used by oracle_test.cpp's compare_value. Built
# unconditionally so the helper has direct unit coverage independent of
# any golden-JSON state.
# ---------------------------------------------------------------------------

add_executable(formulon_oracle_anchor_test
  oracle_anchor.cpp
  oracle_anchor_test.cpp
)

target_link_libraries(formulon_oracle_anchor_test
  PRIVATE
    formulon_static
    gtest_main
)

target_include_directories(formulon_oracle_anchor_test
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
    ${CMAKE_SOURCE_DIR}
)

formulon_apply_warnings(formulon_oracle_anchor_test)

gtest_discover_tests(formulon_oracle_anchor_test)

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

# Register one CTest entry per imported workbook sheet.  A single CTest entry
# for all ~12,500 cases made an allowlist exception capable of suppressing
# every secondary-oracle failure.  Per-case registration would add excessive
# process-launch overhead, while a golden file is a useful failure unit (at
# most 400 cases from one upstream worksheet).
file(GLOB _formulon_ironcalc_goldens CONFIGURE_DEPENDS
  "${CMAKE_CURRENT_SOURCE_DIR}/golden/ironcalc/*.golden.json"
)
foreach(_formulon_ironcalc_golden IN LISTS _formulon_ironcalc_goldens)
  get_filename_component(_formulon_ironcalc_file "${_formulon_ironcalc_golden}" NAME)
  string(REGEX REPLACE "\\.golden\\.json$" "" _formulon_ironcalc_suffix "${_formulon_ironcalc_file}")
  string(REPLACE "__" "_" _formulon_ironcalc_suite_suffix "${_formulon_ironcalc_suffix}")
  set(_formulon_ironcalc_suite "ironcalc_${_formulon_ironcalc_suite_suffix}")
  add_test(
    NAME "formulon_ironcalc_${_formulon_ironcalc_suffix}"
    COMMAND formulon_ironcalc_oracle_tests --gtest_brief=1
            "--gtest_filter=Oracle/OracleTest.Matches/${_formulon_ironcalc_suite}_*"
  )
  set_tests_properties("formulon_ironcalc_${_formulon_ironcalc_suffix}" PROPERTIES LABELS "ironcalc")
endforeach()

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

# `targets.yaml` is the eligibility source of truth.  The JSON marker is a
# required second cross-check, not a way to promote a directory by itself.
# If the optional Python/PyYAML tooling is unavailable, fail closed and keep
# all external captures out of active CTest coverage.
find_package(Python3 COMPONENTS Interpreter QUIET)
set(_formulon_oracle_python "${Python3_EXECUTABLE}")
if(EXISTS "${CMAKE_SOURCE_DIR}/tools/oracle/.venv/bin/python")
  set(_formulon_oracle_python "${CMAKE_SOURCE_DIR}/tools/oracle/.venv/bin/python")
endif()
set(_manifest_active_variants "")
set(_manifest_workbook_active "")

# Whether a golden tree counts as active CTest coverage is decided here, at
# configure time, from files the build does not otherwise depend on. Landing
# a capture therefore had no effect until someone happened to re-run cmake,
# and the failure was silent: `ctest` reported a clean run over the cases it
# still didn't know about. Make the two markers configure inputs so the
# generator picks the change up on the next build instead.
set_property(
  DIRECTORY APPEND PROPERTY CMAKE_CONFIGURE_DEPENDS
  "${CMAKE_SOURCE_DIR}/tools/oracle/targets.yaml"
  "${CMAKE_CURRENT_SOURCE_DIR}/golden_wb/PROVENANCE.json"
)

if(Python3_Interpreter_FOUND)
  execute_process(
    COMMAND ${_formulon_oracle_python} ${CMAKE_SOURCE_DIR}/tools/oracle/provenance.py
            active-variants --repo-root ${CMAKE_SOURCE_DIR}
            --targets-file ${CMAKE_SOURCE_DIR}/tools/oracle/targets.yaml
    OUTPUT_VARIABLE _manifest_active_variants_raw
    OUTPUT_STRIP_TRAILING_WHITESPACE
    ERROR_QUIET
  )
  if(_manifest_active_variants_raw)
    string(REPLACE "\n" ";" _manifest_active_variants "${_manifest_active_variants_raw}")
  endif()
  execute_process(
    COMMAND ${_formulon_oracle_python} ${CMAKE_SOURCE_DIR}/tools/oracle/provenance.py
            workbook-active --repo-root ${CMAKE_SOURCE_DIR}
            --targets-file ${CMAKE_SOURCE_DIR}/tools/oracle/targets.yaml
    OUTPUT_VARIABLE _manifest_workbook_active
    OUTPUT_STRIP_TRAILING_WHITESPACE
    ERROR_QUIET
  )
endif()

# A directory under variants/ is not automatically an active oracle.  In
# particular, the historical win-365-ja_JP capture is deliberately kept for
# research but its manifest status is `wanted` and its provenance record says
# `active_ctest: false`.  Parse the tiny, committed JSON marker here rather
# than letting an old golden tree silently become CI coverage merely because
# FORMULON_ORACLE_VARIANTS=ON was enabled.
function(formulon_oracle_variant_active tag out_var)
  list(FIND _manifest_active_variants "${tag}" _manifest_variant_index)
  if(_manifest_variant_index LESS 0)
    set(${out_var} FALSE PARENT_SCOPE)
    return()
  endif()
  set(_provenance "${CMAKE_CURRENT_SOURCE_DIR}/variants/${tag}/PROVENANCE.json")
  if(NOT EXISTS "${_provenance}")
    set(${out_var} FALSE PARENT_SCOPE)
    return()
  endif()
  file(READ "${_provenance}" _record)
  # Keep the status in the directory marker as a second line of defence:
  # even if a stale active marker is copied over a `status: wanted` target,
  # CTest must not admit it.  provenance.py performs the manifest-level
  # cross-check; this configure-time guard makes the exclusion immediate.
  string(REGEX MATCH "\"status\"[ \t\r\n]*:[ \t\r\n]*\"(primary|scaffolded|wanted)\"" _status "${_record}")
  if(NOT _status OR _record MATCHES "\"status\"[ \t\r\n]*:[ \t\r\n]*\"wanted\"")
    set(${out_var} FALSE PARENT_SCOPE)
    return()
  endif()
  string(REGEX MATCH "\"classification\"[ \t\r\n]*:[ \t\r\n]*\"(active|verified)\"" _classification "${_record}")
  string(REGEX MATCH "\"verified\"[ \t\r\n]*:[ \t\r\n]*true" _verified "${_record}")
  string(REGEX MATCH "\"active_ctest\"[ \t\r\n]*:[ \t\r\n]*true" _active_ctest "${_record}")
  if(_classification AND _verified AND _active_ctest)
    set(${out_var} TRUE PARENT_SCOPE)
  else()
    set(${out_var} FALSE PARENT_SCOPE)
  endif()
endfunction()

if(FORMULON_ORACLE_VARIANTS)
  set(_variant_root "${CMAKE_CURRENT_SOURCE_DIR}/variants")
  set(_variant_dirs_default "")
  if(EXISTS "${_variant_root}")
    file(GLOB _variant_subdirs RELATIVE "${_variant_root}" "${_variant_root}/*")
    foreach(_tag IN LISTS _variant_subdirs)
      set(_golden "${_variant_root}/${_tag}/golden")
      formulon_oracle_variant_active("${_tag}" _variant_active)
      if(_variant_active AND IS_DIRECTORY "${_golden}")
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

  # -------------------------------------------------------------------------
  # Workbook variant oracle target (opt-in)
  #
  # Mirrors formulon_oracle_variant_tests for the workbook track: loads
  # goldens from tests/oracle/variants/<tag>/golden_wb/ (note `golden_wb`,
  # not `golden`) and tags each TEST_P parameter with `__<tag>`. The scan is
  # a parallel pass over the same variants/ root because the workbook track
  # uses a separate golden tree; the discovered list is baked in as a
  # compile-time default and FORMULON_WORKBOOK_ORACLE_VARIANT_DIRS overrides
  # it at runtime. With no variant golden_wb dirs present the binary builds
  # and registers zero cases, so the build stays green.
  # -------------------------------------------------------------------------

  set(_wb_variant_dirs_default "")
  if(EXISTS "${_variant_root}")
    file(GLOB _variant_subdirs RELATIVE "${_variant_root}" "${_variant_root}/*")
    foreach(_tag IN LISTS _variant_subdirs)
      set(_golden_wb "${_variant_root}/${_tag}/golden_wb")
      formulon_oracle_variant_active("${_tag}" _variant_active)
      if(_variant_active AND IS_DIRECTORY "${_golden_wb}")
        if(_wb_variant_dirs_default STREQUAL "")
          set(_wb_variant_dirs_default "${_tag}=${_golden_wb}")
        else()
          set(_wb_variant_dirs_default "${_wb_variant_dirs_default}:${_tag}=${_golden_wb}")
        endif()
      endif()
    endforeach()
  endif()

  add_executable(formulon_workbook_oracle_variant_tests
    json_reader.cpp
    oracle_runner.cpp
    workbook_builder.cpp
    workbook_oracle_runner.cpp
    workbook_oracle_test.cpp
  )

  target_link_libraries(formulon_workbook_oracle_variant_tests
    PRIVATE
      formulon_static
      gtest_main
  )

  target_include_directories(formulon_workbook_oracle_variant_tests
    PRIVATE
      ${CMAKE_SOURCE_DIR}/src
      ${CMAKE_SOURCE_DIR}
  )

  # Empty primary dir so the binary loads ONLY variants.
  # wb_variant_dirs_default gets the configure-time scan; the runtime env
  # var FORMULON_WORKBOOK_ORACLE_VARIANT_DIRS still overrides.
  target_compile_definitions(formulon_workbook_oracle_variant_tests
    PRIVATE
      FORMULON_WORKBOOK_ORACLE_GOLDEN_DIR=""
      FORMULON_WORKBOOK_ORACLE_VARIANT_DIRS_DEFAULT="${_wb_variant_dirs_default}"
  )

  formulon_apply_warnings(formulon_workbook_oracle_variant_tests)

  gtest_discover_tests(formulon_workbook_oracle_variant_tests
    DISCOVERY_MODE PRE_TEST
    PROPERTIES LABELS "VARIANT"
  )
endif()

# ---------------------------------------------------------------------------
# Conditional-formatting (CF) oracle target
#
# Pins Formulon's CF evaluator output for hand-authored scenarios. The
# scenario inputs are built programmatically in the fixtures, but the
# expected match outputs are read from the checked-in golden
# (golden_cf/<suite>.golden.json) so the golden is the single source of
# truth. The paired input artefacts (cases_cf/<suite>.yaml +
# cases_cf/<suite>.case.json) are kept in shape by the Python validator
# `tools/oracle/cf_case_schema.py`.
#
# Registered separately from `formulon_oracle_tests` because the CF
# match data model (DifferentialFormat / ColorScale / DataBar /
# IconSet payloads) is structurally different from the Value-result
# schema the function oracle uses.
# ---------------------------------------------------------------------------

add_executable(formulon_cf_oracle_tests
  cf_oracle_test.cpp
  json_reader.cpp
)

target_link_libraries(formulon_cf_oracle_tests
  PRIVATE
    formulon_static
    gtest_main
)

target_include_directories(formulon_cf_oracle_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
    ${CMAKE_SOURCE_DIR}
)

target_compile_definitions(formulon_cf_oracle_tests
  PRIVATE
    FORMULON_CF_GOLDEN_DIR="${CMAKE_CURRENT_SOURCE_DIR}/golden_cf"
)

formulon_apply_warnings(formulon_cf_oracle_tests)

gtest_discover_tests(formulon_cf_oracle_tests
  PROPERTIES LABELS "oracle"
)

# ---------------------------------------------------------------------------
# Workbook oracle target
#
# Pins Formulon's workbook-level behaviour (pivot tables + print areas)
# against committed golden JSON. Cases are declarative mini-workbook specs
# under cases_wb/<suite>.case.json; the golden JSON under golden_wb/ is
# produced by tools/oracle/workbook_oracle_gen.py on a Windows + Excel
# host. When no golden files exist (the default state before that
# generator has been run) the binary links and runs but registers zero
# parameterized cases, so the build stays green.
#
# Registered separately from formulon_oracle_tests because the workbook
# spec / result data model is structurally different from the Value-result
# schema the function oracle uses. Historical workbook goldens remain
# available for comparison, but provenance gates whether they are active.
# ---------------------------------------------------------------------------

add_executable(formulon_workbook_oracle_tests
  json_reader.cpp
  oracle_runner.cpp
  workbook_builder.cpp
  workbook_oracle_runner.cpp
  workbook_oracle_test.cpp
)

# The checked-in workbook goldens were captured with an unverified Excel
# product/version. Keep them available for historical comparison, but do not
# count them as active Windows 365 CTest coverage until their provenance file
# explicitly opts in. The native builder test below remains active.
set(_workbook_oracle_golden_dir "${CMAKE_CURRENT_SOURCE_DIR}/golden_wb")
set(_workbook_oracle_active FALSE)
set(_workbook_oracle_provenance "${_workbook_oracle_golden_dir}/PROVENANCE.json")
if(EXISTS "${_workbook_oracle_provenance}" AND _manifest_workbook_active STREQUAL "active")
  file(READ "${_workbook_oracle_provenance}" _workbook_record)
  string(REGEX MATCH "\"status\"[ \t\r\n]*:[ \t\r\n]*\"(primary|scaffolded|wanted)\"" _workbook_status "${_workbook_record}")
  string(REGEX MATCH "\"classification\"[ \t\r\n]*:[ \t\r\n]*\"(active|verified)\"" _workbook_classification "${_workbook_record}")
  string(REGEX MATCH "\"verified\"[ \t\r\n]*:[ \t\r\n]*true" _workbook_verified "${_workbook_record}")
  string(REGEX MATCH "\"active_ctest\"[ \t\r\n]*:[ \t\r\n]*true" _workbook_active_ctest "${_workbook_record}")
  if(_workbook_status AND NOT _workbook_record MATCHES "\"status\"[ \t\r\n]*:[ \t\r\n]*\"wanted\"" AND _workbook_classification AND _workbook_verified AND _workbook_active_ctest)
    set(_workbook_oracle_active TRUE)
  endif()
endif()
if(NOT _workbook_oracle_active)
  set(_workbook_oracle_golden_dir "")
endif()

target_link_libraries(formulon_workbook_oracle_tests
  PRIVATE
    formulon_static
    gtest_main
)

target_include_directories(formulon_workbook_oracle_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
    ${CMAKE_SOURCE_DIR}
)

target_compile_definitions(formulon_workbook_oracle_tests
  PRIVATE
    FORMULON_WORKBOOK_ORACLE_GOLDEN_DIR="${_workbook_oracle_golden_dir}"
)

formulon_apply_warnings(formulon_workbook_oracle_tests)

gtest_discover_tests(formulon_workbook_oracle_tests
  DISCOVERY_MODE PRE_TEST
  PROPERTIES LABELS "oracle"
)

# ---------------------------------------------------------------------------
# Workbook builder unit test
#
# Direct gtest for `build_pivot_from_spec`, the test-harness builder that
# turns a declarative workbook + pivot spec into in-memory engine objects.
# This binary is the phase's real local verification while the
# workbook-oracle golden-diff path stays dormant (no Windows-Excel host to
# generate pivot goldens). Built unconditionally and registered under the
# "oracle" CTest label.
# ---------------------------------------------------------------------------

add_executable(formulon_workbook_builder_test
  json_reader.cpp
  oracle_runner.cpp
  workbook_builder.cpp
  workbook_builder_test.cpp
)

target_link_libraries(formulon_workbook_builder_test
  PRIVATE
    formulon_static
    gtest_main
)

target_include_directories(formulon_workbook_builder_test
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
    ${CMAKE_SOURCE_DIR}
)

formulon_apply_warnings(formulon_workbook_builder_test)

gtest_discover_tests(formulon_workbook_builder_test
  PROPERTIES LABELS "oracle"
)

# Metadata checks are read-only and deliberately separate from the executable
# golden tests.  The closure check reports external-pending suites (including
# an incomplete historical print suite) without fabricating or deleting data.
find_package(Python3 COMPONENTS Interpreter QUIET)
if(Python3_Interpreter_FOUND)
  add_test(
    NAME formulon_oracle_provenance_check
    COMMAND ${_formulon_oracle_python} ${CMAKE_SOURCE_DIR}/tools/oracle/provenance.py check
  )
  set_tests_properties(formulon_oracle_provenance_check PROPERTIES LABELS "oracle;oracle-metadata")
  add_test(
    NAME formulon_oracle_policy_unit
    COMMAND ${_formulon_oracle_python} -m unittest tools.oracle.provenance_test
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  set_tests_properties(formulon_oracle_policy_unit PROPERTIES LABELS "oracle;oracle-metadata")
  add_test(
    NAME formulon_macos_excel_formula_unit
    COMMAND ${_formulon_oracle_python} -m unittest tools.oracle.macos_excel_test
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  set_tests_properties(formulon_macos_excel_formula_unit PROPERTIES LABELS "oracle;oracle-metadata")
  option(FORMULON_ORACLE_RELEASE_GATES
    "Register strict external-capture oracle gates (release CI only)"
    OFF
  )
  if(FORMULON_ORACLE_RELEASE_GATES)
    add_test(
      NAME formulon_workbook_closure_check
      COMMAND ${_formulon_oracle_python} ${CMAKE_SOURCE_DIR}/tools/oracle/workbook_closure_check.py --require-active
    )
    # Release workflows opt in explicitly with -DFORMULON_ORACLE_RELEASE_GATES=ON
    # and ctest -L oracle-release.  The strict capture gate is not registered
    # in a normal developer/CI build merely because it has a label.
    set_tests_properties(formulon_workbook_closure_check PROPERTIES LABELS "oracle-release")
  endif()
  add_test(
    NAME formulon_workbook_closure_check_optional
    COMMAND ${_formulon_oracle_python} ${CMAKE_SOURCE_DIR}/tools/oracle/workbook_closure_check.py
  )
  set_tests_properties(formulon_workbook_closure_check_optional PROPERTIES LABELS "oracle;oracle-metadata")
  add_test(
    NAME formulon_workbook_case_schema_negative
    COMMAND ${_formulon_oracle_python} -m unittest tools.oracle.workbook_case_schema_test
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  set_tests_properties(formulon_workbook_case_schema_negative PROPERTIES LABELS "oracle;oracle-metadata")
endif()
