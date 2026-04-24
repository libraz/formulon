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
  unit/sheet_test.cpp
  unit/workbook_test.cpp
  unit/utils/double_format_test.cpp
  unit/deps/miniz_smoke_test.cpp
  unit/deps/pugixml_smoke_test.cpp
  unit/parser/tokenizer_test.cpp
  unit/parser/ast_test.cpp
  unit/parser/ast_dump_test.cpp
  unit/parser/parser_test.cpp
  unit/parser/parser_error_test.cpp
  unit/parser/let_test.cpp
  unit/eval/tree_walker_test.cpp
  unit/eval/eval_context_test.cpp
  unit/eval/eval_state_test.cpp
  unit/eval/cross_sheet_test.cpp
  unit/eval/utf8_length_test.cpp
  unit/eval/function_registry_test.cpp
  unit/eval/builtins_test.cpp
  unit/eval/builtins_logical_test.cpp
  unit/eval/builtins_math_test.cpp
  unit/eval/builtins_math2_test.cpp
  unit/eval/builtins_math3_test.cpp
  unit/eval/builtins_math4_test.cpp
  unit/eval/builtins_text_test.cpp
  unit/eval/builtins_text2_test.cpp
  unit/eval/builtins_textb_test.cpp
  unit/eval/builtins_text3_test.cpp
  unit/eval/builtins_width_test.cpp
  unit/eval/text_format_number_test.cpp
  unit/eval/text_format_date_test.cpp
  unit/eval/value_numbervalue_test.cpp
  unit/eval/builtins_info_test.cpp
  unit/eval/builtins_info_extra_test.cpp
  unit/eval/info_lazy_test.cpp
  unit/eval/builtins_count_test.cpp
  unit/eval/builtins_shape_test.cpp
  unit/eval/builtins_stats_test.cpp
  unit/eval/builtins_stats_exc_test.cpp
  unit/eval/builtins_stats_a_test.cpp
  unit/eval/builtins_descriptive_stats_test.cpp
  unit/eval/builtins_distributions_test.cpp
  unit/eval/builtins_chisq_test.cpp
  unit/eval/builtins_t_test.cpp
  unit/eval/builtins_f_test.cpp
  unit/eval/stats/special_functions_test.cpp
  unit/eval/stats/incomplete_beta_test.cpp
  unit/eval/criteria_test.cpp
  unit/eval/builtins_countif_test.cpp
  unit/eval/builtins_database_test.cpp
  unit/eval/builtins_ifs_test.cpp
  unit/eval/builtins_lookup_test.cpp
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
  unit/eval/builtins_regression_test.cpp
  unit/eval/builtins_hypothesis_test.cpp
  unit/eval/builtins_sumsq_test.cpp
  unit/eval/builtins_sumx_test.cpp
  unit/eval/builtins_series_sum_test.cpp
  unit/eval/builtins_rank_test.cpp
  unit/eval/builtins_distributions_ext2_test.cpp
  unit/eval/builtins_math5_test.cpp
  unit/eval/text_ops_test.cpp
  unit/eval/let_test.cpp
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
add_subdirectory(oracle)
