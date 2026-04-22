# tests/CMakeLists.txt
#
# Fetches Google Test and defines the formulon_unit_tests executable.

include(FetchGoogleTest)

include(GoogleTest)

set(FORMULON_UNIT_TEST_SOURCES
  unit/expected_test.cpp
  unit/error_test.cpp
  unit/status_macros_test.cpp
  unit/structured_log_test.cpp
  unit/strings_test.cpp
  unit/arena_test.cpp
  unit/workbook_test.cpp
  unit/deps/miniz_smoke_test.cpp
  unit/deps/pugixml_smoke_test.cpp
)

add_executable(formulon_unit_tests ${FORMULON_UNIT_TEST_SOURCES})

target_link_libraries(formulon_unit_tests
  PRIVATE
    formulon_static
    gtest_main
)

target_include_directories(formulon_unit_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
)

formulon_apply_warnings(formulon_unit_tests)

gtest_discover_tests(formulon_unit_tests)

add_subdirectory(integration)
