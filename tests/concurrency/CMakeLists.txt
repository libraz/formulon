# tests/concurrency/CMakeLists.txt
#
# Concurrency tests for the parallel scheduler, recalc engine, function
# registry, and iterative solver. Split across three executables so the
# fast `ctest -LE "SLOW|LOAD|TSAN"` invocation skips the slow / sanitizer
# fixtures cleanly.
#
#   * `formulon_concurrency_tests`        — fast scheduler topology tests
#                                           (no extra label).
#   * `formulon_concurrency_stress_tests` — SLOW-labelled stress fixture.
#   * `formulon_concurrency_tsan_tests`   — TSAN-labelled lifetime,
#                                           registry, iterative-solver,
#                                           and engine-internal tests.
#                                           Built into every config so a
#                                           plain `ctest -L TSAN` run can
#                                           pick them up; the dedicated
#                                           TSan build pulls them in via
#                                           `tools/ci/run_tsan.sh`.

# ---- Fast scheduler topology tests ---------------------------------------

add_executable(formulon_concurrency_tests scheduler_test.cpp)

target_link_libraries(formulon_concurrency_tests
  PRIVATE
    formulon_static
    formulon_test_util
    gtest_main
    Threads::Threads
)

target_include_directories(formulon_concurrency_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
)

formulon_apply_warnings(formulon_concurrency_tests)

# Filter out the SLOW fixture. gtest's --gtest_filter syntax supports
# negative-suite matches via `-`; we exclude the `SchedulerSlow.*` group.
gtest_discover_tests(formulon_concurrency_tests
  TEST_FILTER "-SchedulerSlow.*"
)

# ---- SLOW-labelled stress test -------------------------------------------

add_executable(formulon_concurrency_stress_tests scheduler_test.cpp)

target_link_libraries(formulon_concurrency_stress_tests
  PRIVATE
    formulon_static
    gtest_main
    Threads::Threads
)

target_include_directories(formulon_concurrency_stress_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
)

formulon_apply_warnings(formulon_concurrency_stress_tests)

gtest_discover_tests(formulon_concurrency_stress_tests
  TEST_FILTER "SchedulerSlow.*"
  PROPERTIES LABELS "SLOW"
)

# ---- TSAN-labelled lifetime / engine-internal tests ----------------------
#
# These tests target the contract edges that the topology tests do not
# touch: workbook destruction after recalc, concurrent reads on disjoint
# cells during recalc, registry first-use safety, iterative-solver
# thread-safety, and dep-graph read-during-recalc safety.
#
# The label `TSAN` opts every test into the dedicated ThreadSanitizer CI
# job (`tools/ci/run_tsan.sh`) and excludes them from the default fast
# `ctest -LE "SLOW|LOAD|TSAN"` run so they do not slow down the PR loop
# under non-sanitizer builds.

set(FORMULON_CONCURRENCY_TSAN_SOURCES
  recalc_lifetime_test.cpp
  registry_thread_safety_test.cpp
  iterative_solver_thread_test.cpp
  recalc_engine_internal_test.cpp
  recalc_engine_thread_safety_test.cpp
)

add_executable(formulon_concurrency_tsan_tests ${FORMULON_CONCURRENCY_TSAN_SOURCES})

target_link_libraries(formulon_concurrency_tsan_tests
  PRIVATE
    formulon_static
    gtest_main
    Threads::Threads
)

target_include_directories(formulon_concurrency_tsan_tests
  PRIVATE
    ${CMAKE_SOURCE_DIR}/src
)

formulon_apply_warnings(formulon_concurrency_tsan_tests)

gtest_discover_tests(formulon_concurrency_tsan_tests
  PROPERTIES LABELS "TSAN"
)

# Convenience custom target: builds the TSan test executable and prints
# the path to `tools/ci/run_tsan.sh` for the operator. The script itself
# manages the dedicated `build-tsan/` directory; this CMake target only
# guarantees the source-side prerequisites are up to date when invoked
# from the regular `build/` configuration. Wire-through for CI runners
# that want a single CMake target to depend on.
add_custom_target(formulon_tsan_gate
  DEPENDS formulon_concurrency_tsan_tests
  COMMAND ${CMAKE_COMMAND} -E echo
          "TSan gate harness: bash ${CMAKE_SOURCE_DIR}/tools/ci/run_tsan.sh"
  COMMENT "TSan-test sources up to date; run tools/ci/run_tsan.sh to drive the dedicated TSan build"
)
