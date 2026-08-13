# tests/CMakeLists.txt
#
# Fetches Google Test and defines the formulon_unit_tests executable.

include(FetchGoogleTest)

include(GoogleTest)

# Shared test infrastructure (Arena holders, EvalSource wrappers,
# WorkbookBuilder, EXPECT_VALUE_* macros). Declared before the
# individual test executables so they can link against it.
add_subdirectory(util)

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
  unit/utils/checked_mul_test.cpp
  unit/utils/a1_column_test.cpp
  unit/utils/budget_charge_test.cpp
  unit/utils/double_format_test.cpp
  unit/utils/rect_iterator_test.cpp
  unit/utils/resource_budget_test.cpp
  unit/utils/thread_launch_test.cpp
  unit/deps/miniz_smoke_test.cpp
  unit/deps/pugixml_smoke_test.cpp
  unit/io/a1_ref_test.cpp
  unit/io/cell_parser_test.cpp
  unit/io/cf_overlay_test.cpp
  unit/io/cf_reader_test.cpp
  unit/io/cf_writer_test.cpp
  unit/io/comments_roundtrip_test.cpp
  unit/io/data_validation_roundtrip_test.cpp
  unit/io/defined_names_test.cpp
  unit/io/future_functions_test.cpp
  unit/io/hyperlink_roundtrip_test.cpp
  unit/io/iso_date_test.cpp
  unit/io/merge_roundtrip_test.cpp
  unit/io/ooxml_reader_test.cpp
  unit/io/ooxml_unknown_kind_test.cpp
  unit/io/ooxml_writer_cell_test.cpp
  unit/io/ooxml_xlsm_test.cpp
  unit/io/pivot_cache_reader_test.cpp
  unit/io/pivot_cache_writer_test.cpp
  unit/io/pivot_table_reader_test.cpp
  unit/io/pivot_table_writer_test.cpp
  unit/io/sax_shared_formula_test.cpp
  unit/io/sax_xml_reader_test.cpp
  unit/io/sheet_name_validation_test.cpp
  unit/io/sheet_reader_lexing_parity_test.cpp
  unit/io/sheet_reader_test.cpp
  unit/io/sst_reader_test.cpp
  unit/io/storage_prefix_container_test.cpp
  unit/io/styles_reader_test.cpp
  unit/io/styles_writer_test.cpp
  unit/io/tables_reader_test.cpp
  unit/io/workbook_kind_test.cpp
  unit/io/xml_escape_test.cpp
  unit/io/xml_utils_test.cpp
  unit/io/xsd_bool_test.cpp
  unit/io/xsd_double_test.cpp
  unit/io/zip_reader_test.cpp
  unit/io/format_detect_test.cpp
  unit/io/roundtrip_symmetry_test.cpp
  unit/io/pivot_roundtrip_symmetry_test.cpp
  unit/io/sheet_roundtrip_symmetry_test.cpp
  unit/io/xlsb_roundtrip_symmetry_test.cpp
  unit/io/xlsb/func_id_fixture_test.cpp
  unit/io/xlsb/func_id_test.cpp
  unit/io/xlsb/ptg_test.cpp
  unit/io/xlsb/ptg_codec_test.cpp
  unit/io/xlsb/reader_test.cpp
  unit/io/xlsb/record_test.cpp
  unit/io/xlsb/record_writer_test.cpp
  unit/io/xlsb/sheet_tail_test.cpp
  unit/io/xlsb/sst_writer_test.cpp
  unit/io/xlsb/writer_test.cpp
  unit/io/xlsb_fidelity_test.cpp
  unit/parser/tokenizer_test.cpp
  unit/parser/ast_test.cpp
  unit/parser/ast_dump_test.cpp
  unit/parser/ast_format_test.cpp
  unit/parser/ast_shift_test.cpp
  unit/parser/parser_test.cpp
  unit/parser/parser_error_test.cpp
  unit/parser/let_test.cpp
  unit/parser/ref_transforms_test.cpp
  unit/eval/bytecode_test.cpp
  unit/eval/opcode_meta_test.cpp
  unit/eval/jp_kana_table_test.cpp
  unit/eval/compiler_test.cpp
  unit/eval/optimizer_test.cpp
  unit/eval/vm_test.cpp
  unit/eval/vm_parity_test.cpp
  unit/eval/tree_walker_test.cpp
  unit/eval/defined_name_resolve_test.cpp
  unit/eval/eval_sweep_batch_test.cpp
  unit/eval/scalar_ops_test.cpp
  unit/eval/eval_context_test.cpp
  unit/eval/formula_profile_test.cpp
  unit/eval/eval_recursion_depth_test.cpp
  unit/eval/eval_state_test.cpp
  unit/eval/dep_graph_test.cpp
  unit/eval/dep_extractor_test.cpp
  unit/eval/dirty_set_test.cpp
  unit/eval/iterative_progress_test.cpp
  unit/eval/iterative_solver_test.cpp
  unit/eval/partial_recalc_test.cpp
  unit/eval/recalc_engine_test.cpp
  unit/eval/cell_evaluator_test.cpp
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
  unit/eval/array_alloc_test.cpp
  unit/eval/array_producer_test.cpp
  unit/eval/builtins_row_column_array_test.cpp
  unit/eval/sumproduct_array_test.cpp
  unit/eval/eval_spill_dispatch_test.cpp
  unit/eval/eval_context_spill_read_test.cpp
  unit/eval/spill_committer_test.cpp
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
  unit/eval/date_time_1904_test.cpp
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
  unit/eval/lambda_format_test.cpp
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
  unit/pivot/pivot_layout_test.cpp
  unit/pivot/pivot_types_test.cpp
  unit/cf/cf_evaluator_test.cpp
  unit/cf/cf_types_test.cpp
  unit/print/page_setup_test.cpp
  unit/print/print_area_test.cpp
  unit/print/pagination_test.cpp
  unit/registry_catalog_test.cpp
)

add_executable(formulon_unit_tests ${FORMULON_UNIT_TEST_SOURCES})

target_link_libraries(formulon_unit_tests
  PRIVATE
    formulon_static
    formulon_test_util
    gtest_main
)

target_include_directories(formulon_unit_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
    ${CMAKE_CURRENT_SOURCE_DIR}/unit/eval
)

formulon_apply_warnings(formulon_unit_tests)

# Drift-detection catalog path for tests/unit/registry_catalog_test.cpp.
# The test opens this file at runtime and asserts that every registered /
# lazy-dispatched function name appears in it.
target_compile_definitions(formulon_unit_tests
  PRIVATE
    FORMULON_CATALOG_PATH="${CMAKE_SOURCE_DIR}/tools/catalog/functions.txt"
    # Root of the test fixture tree; read at runtime by the read<->write
    # symmetry tests (tests/unit/io/roundtrip_symmetry_test.cpp) to load real
    # Excel-produced packages under tests/fixtures/.
    FORMULON_FIXTURES_DIR="${CMAKE_SOURCE_DIR}/tests/fixtures"
)

# `gtest_discover_tests` registers the suite but cannot apply
# per-case CTest labels at configure time — its discovery script runs
# post-build, so `set_property(TEST ...)` would fail with "test does
# not exist". To keep the slowest zip-bomb / cumulative-cap unit cases
# (`ZipReader.ReadEntryRejectsZipBombSizedEntry`,
# `ZipReader.ReadEntryRejectsCumulativeTotal`,
# `ZipReader.ReadEntryRejectsHighRatioEntry`) out of the fast loop we
# rely on the `-j N` parallelism added to `make test`; they run
# alongside the rest in <2 s wall-clock on a modern workstation.
# Separating them into a dedicated SLOW binary would force the gtest
# fixture to be duplicated, which is a worse trade.
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
  add_test(
    NAME FunctionAvailabilityDrift
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/catalog/status.py --check-c-api-availability
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  # Snapshot test for the binding codegen. Keep the discovery pattern
  # explicit: the JIS generator has its own CTest entry below, so a broad
  # `*_test.py` pattern would register that check twice.
  add_test(
    NAME BindingCodegenDrift
    COMMAND ${_FORMULON_BEHAVIOR_PY} -m unittest discover -s
            ${CMAKE_SOURCE_DIR}/tests/codegen -p "binding_manifest_test.py"
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  # The JIS X 0208 data table is generated by a deterministic Python
  # renderer. Run the explicit unittest file so CTest covers both the tracked
  # snapshot and the temporary drift/missing-fixture negative checks; the
  # standalone generator `--check` is intentionally not registered separately.
  add_test(
    NAME Jis0208TableDrift
    COMMAND ${_FORMULON_BEHAVIOR_PY} -m unittest
            ${CMAKE_SOURCE_DIR}/tests/codegen/jis0208_generated_test.py
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  # Marked SLOW because the YAML-vocabulary scan reads the whole behaviour
  # corpus and is the single largest non-oracle test (~1 s in release).
  # Behaviour drift is checked by `make test-slow` and CI's nightly tier
  # rather than the fast-loop gate.
  set_tests_properties(BehaviorDrift PROPERTIES LABELS "SLOW")

  # Manual-sync file drift guards: cross-checks the C-ABI export list, the
  # WASM/N-API `.d.ts` surfaces, and a README method-count blurb against
  # what is actually registered in the source. Pure static regex analysis
  # (no build artifacts, no network), so it stays in the fast-loop gate.
  add_test(
    NAME binding_drift_python_exports
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/dev/check_binding_drift.py python-exports
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  add_test(
    NAME binding_drift_python_struct_layouts
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/dev/check_binding_drift.py python-struct-layouts
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  add_test(
    NAME binding_drift_dts_wasm
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/dev/check_binding_drift.py dts-wasm
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  add_test(
    NAME binding_drift_dts_node
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/dev/check_binding_drift.py dts-node
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  add_test(
    NAME binding_drift_readme_counts
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/dev/check_binding_drift.py readme-counts
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  add_test(
    NAME binding_drift_dts_enums
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/dev/check_binding_drift.py dts-enums
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  add_test(
    NAME binding_drift_style_record_fields
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/dev/check_binding_drift.py style-record-fields
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  # Structural invariants enforced by "exactly one implementation exists":
  # the bounds-checked array-allocation seam (what keeps a wasm32
  # `rows * cols` from wrapping into a short buffer) and the raw-XML
  # retention helpers (what keeps a consumed part's unmodelled content from
  # vanishing on save). Static regex analysis over src/.
  add_test(
    NAME SeamIntegrityArrayAlloc
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/dev/check_source_seams.py array-alloc
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  add_test(
    NAME SeamIntegrityXmlPassthrough
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/dev/check_source_seams.py xml-passthrough
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  add_test(
    NAME DivergenceRecordIntegrity
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/oracle/divergence_check.py
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  add_test(
    NAME OracleVariantDivergenceIntegrity
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/oracle/divergence_check.py
            --input ${CMAKE_SOURCE_DIR}/tests/oracle/variants/win-365-ja_JP/divergence.yaml
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  add_test(
    NAME OracleGeneratorDivergenceMetadata
    COMMAND ${_FORMULON_BEHAVIOR_PY} -m tools.oracle.oracle_gen_metadata_test
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
  add_test(
    NAME OracleGoldenCoverage
    COMMAND ${_FORMULON_BEHAVIOR_PY} ${CMAKE_SOURCE_DIR}/tools/oracle/golden_coverage_check.py
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  )
endif()

add_subdirectory(integration)
add_subdirectory(concurrency)
add_subdirectory(oracle)
add_subdirectory(c_api)
add_subdirectory(bench)
add_subdirectory(parity)
if(FM_BUILD_FUZZ)
  add_subdirectory(fuzz)
endif()
if(FM_BUILD_CLI)
  add_subdirectory(cli)
endif()
