# tests/integration/CMakeLists.txt
#
# Integration tests exercise multiple subsystems together. Each file under
# this directory is linked into a separate executable so failures stay
# isolated from the fast unit-test suite.

set(FORMULON_INTEGRATION_TEST_SOURCES
  legacy_at_test.cpp
  ooxml_corpus_test.cpp
  ooxml_metadata_test.cpp
  ooxml_minimal_writer_test.cpp
  ooxml_roundtrip_test.cpp
  spill_collision_test.cpp
  table_boundary_test.cpp
  workbook_iterative_test.cpp
  workbook_recalc_test.cpp
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

# ---------------------------------------------------------------------------
# Slow-labelled 100-book OOXML round-trip corpus.
#
# Lives in its own executable so the per-binary fast-suite budget is not
# inflated by 100 two-cycle pipelines (some of which exercise 100x100
# cell shapes). The corresponding ctest cases are tagged with the
# `SLOW` label so the default fast suite (`ctest -LE "SLOW|LOAD"`)
# skips them; opt-in invocations (`ctest -L SLOW` or the full suite)
# pick them up.
# ---------------------------------------------------------------------------

add_executable(formulon_ooxml_corpus_100_tests ooxml_corpus_100_test.cpp)

target_link_libraries(formulon_ooxml_corpus_100_tests
  PRIVATE
    formulon_static
    gtest_main
)

target_include_directories(formulon_ooxml_corpus_100_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
)

formulon_apply_warnings(formulon_ooxml_corpus_100_tests)

# `TEST_LIST` lets us label every discovered ctest entry without naming
# them one-by-one (gtest_discover_tests synthesises one ctest test per
# parameterised case).
gtest_discover_tests(formulon_ooxml_corpus_100_tests
  TEST_LIST formulon_ooxml_corpus_100_test_list
  PROPERTIES LABELS "SLOW"
)

# ---------------------------------------------------------------------------
# SAX vs DOM parity gate over the same 100-book feature matrix.
#
# Drives every book through both the production DOM reader and the
# streaming SAX reader and asserts that the resulting workbooks are
# cell-by-cell identical. SLOW-labelled for the same reasons as the
# DOM-only corpus.
# ---------------------------------------------------------------------------

add_executable(formulon_sax_corpus_parity_tests sax_corpus_parity_test.cpp)

target_link_libraries(formulon_sax_corpus_parity_tests
  PRIVATE
    formulon_static
    gtest_main
)

target_include_directories(formulon_sax_corpus_parity_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
)

formulon_apply_warnings(formulon_sax_corpus_parity_tests)

gtest_discover_tests(formulon_sax_corpus_parity_tests
  TEST_LIST formulon_sax_corpus_parity_test_list
  PROPERTIES LABELS "SLOW"
)
