# tests/parity/CMakeLists.txt
#
# Cross-language parity gate: runs run_parity.py, which evaluates a fixed
# corpus of formulas through the native CLI, the staged npm package, and
# the installed Python wheel, and asserts every available channel agrees
# at %.15g-canonicalized IEEE-754 bit granularity.
#
# The harness reports CTest's conventional SKIP when fewer than two channels
# are available. With two or more, every channel must agree and satisfy each
# fixture's expectation.

if(FM_BUILD_TESTING)
  find_package(Python3 COMPONENTS Interpreter)
  if(Python3_Interpreter_FOUND)
    add_test(
      NAME parity_test
      COMMAND ${Python3_EXECUTABLE} ${CMAKE_SOURCE_DIR}/tests/parity/run_parity.py
      WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
    )
    set_tests_properties(parity_test PROPERTIES LABELS "PARITY")
    set_tests_properties(parity_test PROPERTIES SKIP_RETURN_CODE 77)
    add_test(
      NAME parity_runner_test
      COMMAND ${Python3_EXECUTABLE} ${CMAKE_SOURCE_DIR}/tests/parity/run_parity_test.py
      WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
    )
  endif()
endif()
