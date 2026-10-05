# tests/c_api/CMakeLists.txt
#
# C ABI regression suite. Driven by gtest (compiled as C++) but only
# touches the pure-C surface declared in `src/c_api/formulon_c.h`.

set(FORMULON_C_API_TEST_SOURCES
  formulon_c_adhoc_arena_cap_test.cpp
  formulon_c_calc_mode_test.cpp
  formulon_c_cell_geometry_test.cpp
  formulon_c_cf_authoring_test.cpp
  formulon_c_cf_evaluation_test.cpp
  formulon_c_cf_overlay_test.cpp
  formulon_c_cf_validation_test.cpp
  formulon_c_cf_visual_payload_test.cpp
  formulon_c_clock_test.cpp
  formulon_c_display_text_test.cpp
  formulon_c_evaluate_formula_test.cpp
  formulon_c_external_links_test.cpp
  formulon_c_function_metadata_test.cpp
  formulon_c_lambda_text_test.cpp
  formulon_c_logging_test.cpp
  formulon_c_named_styles_test.cpp
  formulon_c_pivot_authoring_test.cpp
  formulon_c_pivot_filters_test.cpp
  formulon_c_pivot_grouping_test.cpp
  formulon_c_pivot_lifecycle_test.cpp
  formulon_c_pivot_read_test.cpp
  formulon_c_print_settings_test.cpp
  formulon_c_protection_test.cpp
  formulon_c_read_scratch_test.cpp
  formulon_c_sheet_features_test.cpp
  formulon_c_sheet_layout_test.cpp
  formulon_c_spill_test.cpp
  formulon_c_styles_alignment_test.cpp
  formulon_c_styles_cell_xf_test.cpp
  formulon_c_styles_dxf_test.cpp
  formulon_c_styles_fills_borders_test.cpp
  formulon_c_styles_fonts_test.cpp
  formulon_c_styles_lifetime_test.cpp
  formulon_c_styles_mutation_test.cpp
  formulon_c_styles_num_formats_test.cpp
  formulon_c_test.cpp
  formulon_c_values_test.cpp
  formulon_c_save_test.cpp
  formulon_c_trace_test.cpp
  formulon_c_zip_budget_test.cpp
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

# Keep the advertised C11 header contract honest. The primary C API tests
# are C++ gtests, so compile a tiny C translation unit with an actual C
# compiler as a separate CTest entry.
find_program(FORMULON_C_COMPILER NAMES cc clang gcc REQUIRED)
add_test(
  NAME formulon_c_header_compile
  COMMAND ${FORMULON_C_COMPILER} -std=c11 -fsyntax-only -I${CMAKE_SOURCE_DIR}/src
          ${CMAKE_CURRENT_SOURCE_DIR}/formulon_c_header_compile.c
)
