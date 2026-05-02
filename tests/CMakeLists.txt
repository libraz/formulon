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
  unit/value_test.cpp
  unit/value_array_test.cpp
  unit/sheet_test.cpp
  unit/sheet_spill_test.cpp
  unit/sheet_cached_text_test.cpp
  unit/workbook_test.cpp
  unit/utils/double_format_test.cpp
  unit/deps/miniz_smoke_test.cpp
  unit/deps/pugixml_smoke_test.cpp
  unit/io/cell_parser_test.cpp
  unit/io/defined_names_test.cpp
  unit/io/ooxml_reader_test.cpp
  unit/io/ooxml_unknown_kind_test.cpp
  unit/io/ooxml_writer_cell_test.cpp
  unit/io/ooxml_xlsm_test.cpp
  unit/io/pivot_cache_reader_test.cpp
  unit/io/pivot_cache_writer_test.cpp
  unit/io/pivot_table_reader_test.cpp
  unit/io/pivot_table_writer_test.cpp
  unit/io/sax_xml_reader_test.cpp
  unit/io/sheet_reader_test.cpp
  unit/io/sst_reader_test.cpp
  unit/io/styles_reader_test.cpp
  unit/io/tables_reader_test.cpp
  unit/io/workbook_kind_test.cpp
  unit/io/zip_reader_test.cpp
  unit/io/xlsb/func_id_test.cpp
  unit/io/xlsb/ptg_test.cpp
  unit/io/xlsb/reader_test.cpp
  unit/io/xlsb/record_test.cpp
  unit/io/xlsb/record_writer_test.cpp
  unit/io/xlsb/sst_writer_test.cpp
  unit/io/xlsb/writer_test.cpp
  unit/parser/tokenizer_test.cpp
  unit/parser/ast_test.cpp
  unit/parser/ast_dump_test.cpp
  unit/parser/parser_test.cpp
  unit/parser/parser_error_test.cpp
  unit/parser/let_test.cpp
  unit/eval/bytecode_test.cpp
  unit/eval/compiler_test.cpp
  unit/eval/optimizer_test.cpp
  unit/eval/vm_test.cpp
  unit/eval/vm_parity_test.cpp
  unit/eval/tree_walker_test.cpp
  unit/eval/eval_context_test.cpp
  unit/eval/eval_state_test.cpp
  unit/eval/dep_graph_test.cpp
  unit/eval/dep_extractor_test.cpp
  unit/eval/dirty_set_test.cpp
  unit/eval/iterative_solver_test.cpp
  unit/eval/recalc_engine_test.cpp
  unit/eval/volatile_tracker_test.cpp
  unit/eval/cross_sheet_test.cpp
  unit/eval/utf8_length_test.cpp
  unit/eval/function_registry_test.cpp
  unit/eval/builtins_test.cpp
  unit/eval/builtins_logical_test.cpp
  unit/eval/builtins_floor_ceiling_test.cpp
  unit/eval/builtins_math_test.cpp
  unit/eval/builtins_math2_test.cpp
  unit/eval/builtins_math3_test.cpp
  unit/eval/builtins_math4_test.cpp
  unit/eval/builtins_text_test.cpp
  unit/eval/builtins_text2_test.cpp
  unit/eval/builtins_bahttext_test.cpp
  unit/eval/builtins_textb_test.cpp
  unit/eval/builtins_text3_test.cpp
  unit/eval/builtins_width_test.cpp
  unit/eval/text_format_number_test.cpp
  unit/eval/text_format_date_test.cpp
  unit/eval/value_numbervalue_test.cpp
  unit/eval/builtins_info_test.cpp
  unit/eval/builtins_info_extra_test.cpp
  unit/eval/info_lazy_test.cpp
  unit/eval/phonetic_lazy_test.cpp
  unit/eval/getpivotdata_lazy_test.cpp
  unit/eval/builtins_count_test.cpp
  unit/eval/builtins_shape_test.cpp
  unit/eval/array_producer_test.cpp
  unit/eval/builtins_row_column_array_test.cpp
  unit/eval/sumproduct_array_test.cpp
  unit/eval/eval_spill_dispatch_test.cpp
  unit/eval/eval_context_spill_read_test.cpp
  unit/eval/builtins_stats_test.cpp
  unit/eval/builtins_subtotal_test.cpp
  unit/eval/builtins_aggregate_fn_test.cpp
  unit/eval/builtins_stats_exc_test.cpp
  unit/eval/builtins_stats_a_test.cpp
  unit/eval/builtins_descriptive_stats_test.cpp
  unit/eval/builtins_distributions_test.cpp
  unit/eval/builtins_chisq_test.cpp
  unit/eval/builtins_t_test.cpp
  unit/eval/builtins_f_test.cpp
  unit/eval/stats/special_functions_test.cpp
  unit/eval/stats/incomplete_beta_test.cpp
  unit/eval/coerce_test.cpp
  unit/eval/criteria_test.cpp
  unit/eval/jp_fold_test.cpp
  unit/eval/jis0208_table_test.cpp
  unit/eval/builtins_countif_test.cpp
  unit/eval/builtins_database_test.cpp
  unit/eval/builtins_groupby_test.cpp
  unit/eval/builtins_pivotby_test.cpp
  unit/eval/builtins_ifs_test.cpp
  unit/eval/builtins_lookup_test.cpp
  unit/eval/builtins_implicit_intersection_test.cpp
  unit/eval/builtins_references_test.cpp
  unit/eval/reference_lazy_test.cpp
  unit/eval/builtins_vlookup_test.cpp
  unit/eval/builtins_xlookup_test.cpp
  unit/eval/builtins_datetime_test.cpp
  unit/eval/builtins_datevalue_timevalue_test.cpp
  unit/eval/builtins_financial_test.cpp
  unit/eval/builtins_financial2_test.cpp
  unit/eval/builtins_financial_rates_test.cpp
  unit/eval/builtins_financial_coupon_test.cpp
  unit/eval/builtins_financial_accrual_test.cpp
  unit/eval/builtins_financial_bond_simple_test.cpp
  unit/eval/builtins_financial_duration_test.cpp
  unit/eval/builtins_financial_oddf_test.cpp
  unit/eval/builtins_financial_oddl_test.cpp
  unit/eval/builtins_financial_price_test.cpp
  unit/eval/builtins_financial_yield_test.cpp
  unit/eval/builtins_forecast_ets_test.cpp
  unit/eval/builtins_regression_test.cpp
  unit/eval/builtins_hypothesis_test.cpp
  unit/eval/builtins_sumsq_test.cpp
  unit/eval/builtins_percentof_test.cpp
  unit/eval/builtins_sumx_test.cpp
  unit/eval/builtins_series_sum_test.cpp
  unit/eval/builtins_rank_test.cpp
  unit/eval/builtins_regex_test.cpp
  unit/eval/builtins_distributions_ext2_test.cpp
  unit/eval/builtins_math5_test.cpp
  unit/eval/text_ops_test.cpp
  unit/eval/let_test.cpp
  unit/eval/lambda_test.cpp
  unit/eval/lambda_helpers_test.cpp
  unit/eval/builtins_misc_test.cpp
  unit/eval/builtins_rand_test.cpp
  unit/eval/builtins_engineering_test.cpp
  unit/eval/builtins_engineering_convert_test.cpp
  unit/eval/builtins_engineering_special_test.cpp
  unit/eval/builtins_complex_test.cpp
  unit/eval/builtins_web_test.cpp
  unit/eval/builtins_cube_test.cpp
  unit/eval/builtins_xirr_xnpv_test.cpp
  unit/eval/builtins_service_stubs_test.cpp
  unit/eval/cell_eval_test.cpp
  unit/eval/builtins/dynamic_array_test.cpp
  unit/eval/spill_ref_eval_test.cpp
  unit/eval/filter_eval_test.cpp
  unit/eval/unique_eval_test.cpp
  unit/eval/sort_eval_test.cpp
  unit/eval/sortby_eval_test.cpp
  unit/eval/stack_eval_test.cpp
  unit/eval/choose_axis_eval_test.cpp
  unit/eval/take_drop_eval_test.cpp
  unit/eval/array_reshape_eval_test.cpp
  unit/eval/matrix_ops_eval_test.cpp
  unit/eval/frequency_eval_test.cpp
  unit/eval/linest_eval_test.cpp
  unit/eval/trend_eval_test.cpp
  unit/eval/logest_growth_eval_test.cpp
  unit/eval/textsplit_eval_test.cpp
  unit/eval/trimrange_eval_test.cpp
  unit/parser/spill_ref_test.cpp
  unit/parser/structured_ref_test.cpp
  unit/eval/structured_ref_eval_test.cpp
  unit/pivot/pivot_evaluator_test.cpp
  unit/pivot/pivot_types_test.cpp
  unit/cf/cf_types_test.cpp
  unit/registry_catalog_test.cpp
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

# Drift-detection catalog path for tests/unit/registry_catalog_test.cpp.
# The test opens this file at runtime and asserts that every registered /
# lazy-dispatched function name appears in it.
target_compile_definitions(formulon_unit_tests
  PRIVATE
    FORMULON_CATALOG_PATH="${CMAKE_SOURCE_DIR}/tools/catalog/functions.txt"
)

gtest_discover_tests(formulon_unit_tests)

# Behavior-vocabulary drift check. Parallel to RegistryCatalog.AllRegisteredInCatalog
# but for function sub-behaviors (format codes, era strings, locale separators).
# Prefers the oracle venv Python (which ships PyYAML) over the system interpreter
# so `ctest` actually enforces the invariant on dev machines. Without PyYAML the
# script soft-skips, which keeps CI green on hosts that haven't bootstrapped the
# oracle venv but also means the invariant is not enforced there — acceptable
# because the enforced run happens on any machine with `make oracle-setup`.
set(_FORMULON_ORACLE_VENV_PY "${CMAKE_SOURCE_DIR}/tools/oracle/.venv/bin/python3")
if(EXISTS "${_FORMULON_ORACLE_VENV_PY}")
  set(_FORMULON_BEHAVIOR_PY "${_FORMULON_ORACLE_VENV_PY}")
else()
  find_package(Python3 COMPONENTS Interpreter)
  if(Python3_Interpreter_FOUND)
    set(_FORMULON_BEHAVIOR_PY "${Python3_EXECUTABLE}")
  endif()
endif()
if(DEFINED _FORMULON_BEHAVIOR_PY)
  add_test(
    NAME BehaviorDrift
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/catalog/behaviors.py --check
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
endif()

add_subdirectory(integration)
add_subdirectory(concurrency)
add_subdirectory(oracle)
add_subdirectory(c_api)
add_subdirectory(bench)
add_subdirectory(parity)
if(FM_BUILD_CLI)
  add_subdirectory(cli)
endif()
