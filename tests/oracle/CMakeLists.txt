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
