# tests/integration/CMakeLists.txt
#
# Integration tests exercise multiple subsystems together. Each file under
# this directory is linked into a separate executable so failures stay
# isolated from the fast unit-test suite.

set(FORMULON_INTEGRATION_TEST_SOURCES
  ooxml_minimal_writer_test.cpp
)

add_executable(formulon_integration_tests ${FORMULON_INTEGRATION_TEST_SOURCES})

target_link_libraries(formulon_integration_tests
  PRIVATE
    formulon_static
    gtest_main
)

target_include_directories(formulon_integration_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
)

formulon_apply_warnings(formulon_integration_tests)

gtest_discover_tests(formulon_integration_tests)
