# tests/concurrency/CMakeLists.txt
#
# Concurrency tests for the parallel scheduler. The fast cases live in
# `formulon_concurrency_tests`; the SLOW-labelled stress case lives in
# `formulon_concurrency_stress_tests` so the fast `ctest -LE "SLOW|LOAD"`
# invocation skips it without per-case property gymnastics.

add_executable(formulon_concurrency_tests scheduler_test.cpp)

target_link_libraries(formulon_concurrency_tests
  PRIVATE
    formulon_static
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

# Stress-test variant: same source, gtest filter limited to the slow
# fixture, registered with the SLOW label.
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
