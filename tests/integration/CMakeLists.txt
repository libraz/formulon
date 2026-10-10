# tests/integration/CMakeLists.txt
#
# Integration tests exercise multiple subsystems together. Each file under
# this directory is linked into a separate executable so failures stay
# isolated from the fast unit-test suite.

set(FORMULON_INTEGRATION_TEST_SOURCES
  cf_databar_union_sqref_fixture_test.cpp
  external_links_roundtrip_test.cpp
  formula_corpus_fixture_smoke_test.cpp
  legacy_at_test.cpp
  ooxml_cf_test.cpp
  ooxml_corpus_test.cpp
  ooxml_corpus_cases.cpp
  ooxml_corpus_package_support.cpp
  ooxml_default_passthrough_test.cpp
  ooxml_metadata_test.cpp
  ooxml_minimal_writer_test.cpp
  ooxml_passthrough_rels_test.cpp
  ooxml_phonetic_roundtrip_test.cpp
  ooxml_pivot_test.cpp
  ooxml_print_authoring_test.cpp
  ooxml_pivot_names_roundtrip_test.cpp
  ooxml_roundtrip_test.cpp
  ooxml_sax_metadata_test.cpp
  ooxml_sheet_features_batch_test.cpp
  ooxml_workbook_pr_roundtrip_test.cpp
  ooxml_writer_pivot_test.cpp
  pivot_blank_item_fixture_test.cpp
  pivot_caption_filter_fixture_test.cpp
  pivot_layout_fixture_test.cpp
  pivot_recurring_period_filter_fixture_test.cpp
  pivot_value_date_filter_fixture_test.cpp
  real_excel_fixture_smoke_test.cpp
  recalc_text_lifetime_test.cpp
  xlsb_default_passthrough_test.cpp
  sheet_features_roundtrip_test.cpp
  sheet_layout_roundtrip_test.cpp
  sheet_print_geometry_roundtrip_test.cpp
  sheet_ops_name_rename_test.cpp
  sheet_ops_remove_move_defined_names_test.cpp
  sheet_ops_row_col_edits_test.cpp
  sheet_ops_3d_structural_recalc_test.cpp
  structural_edit_matrix_test.cpp
  spill_collision_test.cpp
  styles_roundtrip_test.cpp
  table_boundary_test.cpp
  workbook_iterative_test.cpp
  workbook_recalc_test.cpp
  zip_load_budget_test.cpp
)

add_executable(formulon_integration_tests ${FORMULON_INTEGRATION_TEST_SOURCES})

target_link_libraries(formulon_integration_tests
  PRIVATE
    formulon_static
    formulon_test_util
    gtest_main
)

target_include_directories(formulon_integration_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
)

# Root of the test fixture tree; mirrors the same define on
# formulon_unit_tests (tests/CMakeLists.txt) so integration tests can load
# real Excel-produced packages under tests/fixtures/ too.
target_compile_definitions(formulon_integration_tests
  PRIVATE
    FORMULON_FIXTURES_DIR="${CMAKE_SOURCE_DIR}/tests/fixtures"
)

formulon_apply_warnings(formulon_integration_tests)

gtest_discover_tests(formulon_integration_tests)

macro(formulon_add_slow_integration_test target source test_list)
  add_executable(${target} ${source})

  target_link_libraries(${target}
    PRIVATE
      formulon_static
      gtest_main
  )

  target_include_directories(${target}
    PRIVATE
      ${CMAKE_SOURCE_DIR}/src
  )

  formulon_apply_warnings(${target})

  gtest_discover_tests(${target}
    TEST_LIST ${test_list}
    PROPERTIES LABELS "SLOW"
  )
endmacro()

# ---------------------------------------------------------------------------
# Slow-labelled 100-book OOXML round-trip corpus.
#
# Lives in its own executable so the per-binary fast-suite budget is not
# inflated by 100 two-cycle pipelines (some of which exercise 100x100
# cell shapes). The corresponding ctest cases are tagged with the
# `SLOW` label so the default fast suite (`ctest -LE "SLOW|BENCH|TSAN"`)
# skips them; opt-in invocations (`ctest -L SLOW` or the full suite)
# pick them up.
# ---------------------------------------------------------------------------

# `TEST_LIST` lets us label every discovered ctest entry without naming
# them one-by-one (gtest_discover_tests synthesises one ctest test per
# parameterised case).
formulon_add_slow_integration_test(
  formulon_ooxml_corpus_100_tests
  ooxml_corpus_100_test.cpp
  formulon_ooxml_corpus_100_test_list
)

# ---------------------------------------------------------------------------
# Slow-labelled extended OOXML round-trip corpus.
#
# Sibling to the 100-book corpus above. Where that corpus pins the
# general round-trip surface (cells / formulas / defined names / tables /
# passthrough) but explicitly excludes styles, conditional formats, and
# pivots, this 48-book corpus extends the same two-cycle pipeline to
# cover those auxiliary subsystems. Lives in its own executable for the
# same per-binary fast-suite-budget reasons; tagged `SLOW`.
# ---------------------------------------------------------------------------

formulon_add_slow_integration_test(
  formulon_ooxml_corpus_extended_tests
  ooxml_corpus_extended_test.cpp
  formulon_ooxml_corpus_extended_test_list
)

# ---------------------------------------------------------------------------
# SAX vs DOM parity gate over the same 100-book feature matrix.
#
# Drives every book through both the production DOM reader and the
# streaming SAX reader and asserts that the resulting workbooks are
# cell-by-cell identical. SLOW-labelled for the same reasons as the
# DOM-only corpus.
# ---------------------------------------------------------------------------

formulon_add_slow_integration_test(
  formulon_sax_corpus_parity_tests
  sax_corpus_parity_test.cpp
  formulon_sax_corpus_parity_test_list
)

# ---------------------------------------------------------------------------
# Slow-labelled recalc registration-scale guard.
#
# Registers a thousand lookup formulas over a 300,000-cell rectangle and
# asserts the dependency graph stays proportional to the formulas rather
# than to the referenced area. Own executable so the fast tier's per-case
# budget is untouched.
# ---------------------------------------------------------------------------

formulon_add_slow_integration_test(
  formulon_recalc_scale_tests
  recalc_scale_test.cpp
  formulon_recalc_scale_test_list
)
