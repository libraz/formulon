# tests/c_api/CMakeLists.txt
#
# C ABI regression suite. Driven by gtest (compiled as C++) but only
# touches the pure-C surface declared in `src/c_api/formulon_c.h`.

set(FORMULON_C_API_TEST_SOURCES
  formulon_c_adhoc_arena_cap_test.cpp
  formulon_c_api_usability_test.cpp
  formulon_c_argument_range_test.cpp
  formulon_c_argument_messages_test.cpp
  formulon_c_auto_calc_test.cpp
  formulon_c_auto_filter_test.cpp
  formulon_c_calc_mode_test.cpp
  formulon_c_cell_stale_test.cpp
  formulon_c_circular_reference_test.cpp
  formulon_c_cell_geometry_test.cpp
  formulon_c_cf_authoring_test.cpp
  formulon_c_cf_evaluation_test.cpp
  formulon_c_cf_overlay_test.cpp
  formulon_c_cf_validation_test.cpp
  formulon_c_cf_visual_payload_test.cpp
  formulon_c_chart_patch_test.cpp
  formulon_c_chart_test.cpp
  formulon_c_clock_test.cpp
  formulon_c_collation_test.cpp
  formulon_c_date_serial_test.cpp
  formulon_c_display_text_test.cpp
  formulon_c_drawing_test.cpp
  formulon_c_evaluate_formula_test.cpp
  formulon_c_external_links_test.cpp
  formulon_c_formula_entry_test.cpp
  formulon_c_entry_r1c1_test.cpp
  formulon_c_function_metadata_test.cpp
  formulon_c_input_records_test.cpp
  formulon_c_hyperlink_style_test.cpp
  formulon_c_input_validation_test.cpp
  formulon_c_lambda_text_test.cpp
  formulon_c_locale_test.cpp
  formulon_c_location_test.cpp
  formulon_c_logging_test.cpp
  formulon_c_named_styles_test.cpp
  formulon_c_natural_key_test.cpp
  formulon_c_parse_error_test.cpp
  formulon_c_pivot_authoring_test.cpp
  formulon_c_pivot_options_test.cpp
  formulon_c_pivot_patch_test.cpp
  formulon_c_pivot_calc_test.cpp
  formulon_c_pivot_definition_test.cpp
  formulon_c_pivot_filters_test.cpp
  formulon_c_pivot_filter_persistence_test.cpp
  formulon_c_pivot_grouping_test.cpp
  formulon_c_pivot_lifecycle_test.cpp
  formulon_c_pivot_read_test.cpp
  formulon_c_pivot_refresh_test.cpp
  formulon_c_pivot_report_test.cpp
  formulon_c_pivot_save_shape_test.cpp
  formulon_c_pivot_blocked_test.cpp
  formulon_c_pivot_shared_item_limit_test.cpp
  formulon_c_print_settings_test.cpp
  formulon_c_protection_test.cpp
  formulon_c_range_style_io_test.cpp
  formulon_c_range_style_runs_test.cpp
  formulon_c_range_style_test.cpp
  formulon_c_read_scratch_test.cpp
  formulon_c_reference_test.cpp
  formulon_c_resource_limit_test.cpp
  formulon_c_sheet_features_test.cpp
  formulon_c_sheet_index_status_test.cpp
  formulon_c_sheet_layout_test.cpp
  formulon_c_spill_test.cpp
  formulon_c_spill_cause_test.cpp
  formulon_c_status_test.cpp
  formulon_c_structured_ref_test.cpp
  formulon_c_array_formula_test.cpp
  formulon_c_styles_alignment_test.cpp
  formulon_c_styles_cell_xf_test.cpp
  formulon_c_styles_dxf_test.cpp
  formulon_c_styles_fills_borders_test.cpp
  formulon_c_styles_fonts_test.cpp
  formulon_c_styles_lifetime_test.cpp
  formulon_c_styles_mutation_test.cpp
  formulon_c_styles_num_formats_test.cpp
  formulon_c_table_test.cpp
  formulon_c_table_toggle_test.cpp
  formulon_c_table_auto_filter_test.cpp
  formulon_c_table_insert_test.cpp
  formulon_c_tab_color_test.cpp
  formulon_c_test.cpp
  formulon_c_theme_test.cpp
  formulon_c_threaded_comments_test.cpp
  formulon_c_view_style_fidelity_test.cpp
  formulon_c_workbook_view_test.cpp
  formulon_c_validation_test.cpp
  formulon_c_workbook_protection_test.cpp
  formulon_c_values_test.cpp
  formulon_c_save_test.cpp
  formulon_c_shapes_sparklines_test.cpp
  formulon_c_data_table_test.cpp
  formulon_c_slicers_test.cpp
  formulon_c_trace_test.cpp
  formulon_c_utf8_input_test.cpp
  formulon_c_zip_budget_test.cpp
  empty_load_test.cpp
  error_text_test.cpp
  index_out_of_range_test.cpp
  partial_recalc_test.cpp
  pivot_field_order_test.cpp
  remove_missing_key_test.cpp
  sheet_index_by_name_test.cpp
  version_number_test.cpp
  workbook_copy_sheet_test.cpp
  workbook_sheet_ops_test.cpp
  formulon_c_insert_delete_cells_test.cpp
  formulon_c_name_identity_test.cpp
)

# The sheet-index status test calls every sheet-scoped function through a
# table generated from the header and the name ledger (PyYAML required).
if(NOT DEFINED _FORMULON_BEHAVIOR_PY)
  find_package(Python3 COMPONENTS Interpreter REQUIRED)
  set(_FORMULON_BEHAVIOR_PY "${Python3_EXECUTABLE}")
endif()
set(_FM_SHEET_INDEX_CALLS_SCRIPT "${CMAKE_SOURCE_DIR}/tools/codegen/gen_sheet_index_calls.py")
set(_FM_SHEET_INDEX_CALLS_DIR "${CMAKE_CURRENT_BINARY_DIR}/generated")
set(_FM_SHEET_INDEX_CALLS_INC "${_FM_SHEET_INDEX_CALLS_DIR}/sheet_index_calls.inc")
add_custom_command(
  OUTPUT "${_FM_SHEET_INDEX_CALLS_INC}"
  COMMAND ${_FORMULON_BEHAVIOR_PY} "${_FM_SHEET_INDEX_CALLS_SCRIPT}" --output "${_FM_SHEET_INDEX_CALLS_INC}"
  DEPENDS
    "${_FM_SHEET_INDEX_CALLS_SCRIPT}"
    "${CMAKE_SOURCE_DIR}/src/c_api/formulon_c.h"
    "${CMAKE_SOURCE_DIR}/tools/api/names.yaml"
    "${CMAKE_SOURCE_DIR}/tools/dev/binding_drift/api_conventions.py"
    "${CMAKE_SOURCE_DIR}/tools/dev/binding_drift/c_header.py"
  COMMENT "Generating the sheet-scoped C ABI call table"
  VERBATIM
)
set_source_files_properties(formulon_c_sheet_index_status_test.cpp
  PROPERTIES OBJECT_DEPENDS "${_FM_SHEET_INDEX_CALLS_INC}")

add_executable(formulon_c_api_tests ${FORMULON_C_API_TEST_SOURCES} "${_FM_SHEET_INDEX_CALLS_INC}")

target_link_libraries(formulon_c_api_tests
  PRIVATE
    formulon_static
    gtest_main
)

target_include_directories(formulon_c_api_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
    # `support/json_cases.h`: the case tables under tools/api/.
    ${CMAKE_SOURCE_DIR}/tests
    ${_FM_SHEET_INDEX_CALLS_DIR}
)

# Excel-saved packages under tests/fixtures/ (the range-style IO tests).
target_compile_definitions(formulon_c_api_tests
  PRIVATE
    FORMULON_FIXTURES_DIR="${CMAKE_SOURCE_DIR}/tests/fixtures"
)

formulon_apply_warnings(formulon_c_api_tests)

gtest_discover_tests(formulon_c_api_tests)

# Range-style scale: a 2.6M-cell styled rect through save and load in both
# formats, with a peak-RSS bound. Its own executable so `ru_maxrss` measures
# this run alone; every case is SLOW.
add_executable(formulon_c_range_style_scale_tests formulon_c_range_style_scale_test.cpp)
target_link_libraries(formulon_c_range_style_scale_tests
  PRIVATE
    formulon_static
    gtest_main
)
target_include_directories(formulon_c_range_style_scale_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
)
formulon_apply_warnings(formulon_c_range_style_scale_tests)
gtest_discover_tests(formulon_c_range_style_scale_tests
  PROPERTIES LABELS "SLOW" TIMEOUT 120
  DISCOVERY_TIMEOUT 120
)

# Keep the advertised C11 header contract honest. The primary C API tests
# are C++ gtests, so compile a tiny C translation unit with an actual C
# compiler as a separate CTest entry.
find_program(FORMULON_C_COMPILER NAMES cc clang gcc REQUIRED)
add_test(
  NAME formulon_c_header_compile
  COMMAND ${FORMULON_C_COMPILER} -std=c11 -fsyntax-only -I${CMAKE_SOURCE_DIR}/src
          ${CMAKE_CURRENT_SOURCE_DIR}/formulon_c_header_compile.c
)
