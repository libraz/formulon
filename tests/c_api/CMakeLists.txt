# tests/c_api/CMakeLists.txt
#
# C ABI regression suite. Driven by gtest (compiled as C++) but only
# touches the pure-C surface declared in `src/c_api/formulon_c.h`.

set(FORMULON_C_API_TEST_SOURCES
  formulon_c_cf_test.cpp
  formulon_c_sheet_features_test.cpp
  formulon_c_sheet_layout_test.cpp
  formulon_c_styles_test.cpp
  formulon_c_test.cpp
  partial_recalc_test.cpp
  workbook_sheet_ops_test.cpp
)

add_executable(formulon_c_api_tests ${FORMULON_C_API_TEST_SOURCES})

target_link_libraries(formulon_c_api_tests
  PRIVATE
    formulon_static
    gtest_main
)

target_include_directories(formulon_c_api_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
)

formulon_apply_warnings(formulon_c_api_tests)

gtest_discover_tests(formulon_c_api_tests)
