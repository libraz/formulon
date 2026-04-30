# tests/parity/CMakeLists.txt
#
# Cross-language parity gate: runs run_parity.py, which evaluates a fixed
# corpus of formulas through the native CLI, the staged npm package, and
# the installed Python wheel, and asserts every available channel agrees
# at %.15g-canonicalized IEEE-754 bit granularity.
#
# The harness is skip-aware -- a missing channel is reported but does not
# fail. It only fails when two or more channels disagreed on at least one
# fixture, so this ctest entry is safe to include even when only the
# native build has been produced.

if(FM_BUILD_TESTING)
  find_package(Python3 COMPONENTS Interpreter)
  if(Python3_Interpreter_FOUND)
    add_test(
      NAME parity_test
      COMMAND ${Python3_EXECUTABLE} ${CMAKE_SOURCE_DIR}/tests/parity/run_parity.py
      WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
    )
    set_tests_properties(parity_test PROPERTIES LABELS "PARITY")
  endif()
endif()
