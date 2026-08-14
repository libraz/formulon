# tests/cli/CMakeLists.txt
#
# `formulon_cli` regression suite. The harness drives the CLI binary
# as a subprocess via `popen` so the tests exercise the same exit-code
# / stdio path real users see. The binary location is injected via
# `FORMULON_CLI_PATH` so the tests do not depend on `$PATH`.

set(FORMULON_CLI_TEST_SOURCES
  cli_test.cpp
)

add_executable(formulon_cli_tests ${FORMULON_CLI_TEST_SOURCES})

target_link_libraries(formulon_cli_tests
  PRIVATE
    formulon_static
    gtest_main
)

target_include_directories(formulon_cli_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
    # `support/ooxml_package_fixture.h`: the reader-side warnings the CLI
    # prints need a package the engine's own writer would never emit.
    ${CMAKE_SOURCE_DIR}/tests
)

# Generator expression resolves to the absolute path of the binary
# `formulon_cli` produces — works under both single- and multi-config
# generators. Ensures the dependency is built before the tests run.
target_compile_definitions(formulon_cli_tests
  PRIVATE
    FORMULON_CLI_PATH="$<TARGET_FILE:formulon_cli>"
)
add_dependencies(formulon_cli_tests formulon_cli)

formulon_apply_warnings(formulon_cli_tests)

gtest_discover_tests(formulon_cli_tests)
