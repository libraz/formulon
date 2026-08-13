# tests/util/CMakeLists.txt
#
# Shared test infrastructure: per-test Arena helpers, evaluator wrappers,
# Value-shape macros, and a fluent WorkbookBuilder. Consolidates the 248
# open-coded `thread_local Arena` setup blocks and 40+ per-file
# `EvalSource()` redefinitions that previously lived across
# `tests/unit/eval/*_test.cpp`.
#
# The library links PUBLIC against `formulon_static` so any test target
# linking `formulon_test_util` gets the production headers (Workbook,
# Sheet, eval::EvalContext, parser, etc.) for free and does not have to
# repeat the include-dir wiring.

add_library(formulon_test_util STATIC
  test_arena.cpp
  test_eval_helpers.cpp
  test_workbook_builder.cpp
)

target_link_libraries(formulon_test_util
  PUBLIC
    formulon_static
    gtest
)

target_include_directories(formulon_test_util
  PUBLIC
    ${CMAKE_SOURCE_DIR}/tests
    ${CMAKE_SOURCE_DIR}/src
)

formulon_apply_warnings(formulon_test_util)

# ---------------------------------------------------------------------------
# Self-tests for the shared test utilities.
#
# Lives in its own executable so a breakage in the helpers is reported
# under a stand-alone ctest entry rather than bundled into the much
# larger formulon_unit_tests binary. Discovered into ctest with the
# default fast-loop label so `ctest -LE "SLOW|BENCH|TSAN"` covers it.
# ---------------------------------------------------------------------------
add_executable(formulon_test_util_tests test_util_test.cpp)

target_link_libraries(formulon_test_util_tests
  PRIVATE
    formulon_test_util
    gtest_main
)

formulon_apply_warnings(formulon_test_util_tests)

gtest_discover_tests(formulon_test_util_tests)
