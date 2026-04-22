# FetchGoogleTest.cmake
#
# Downloads and configures GoogleTest v1.14.0 using FetchContent.
# Only included from tests/CMakeLists.txt when FM_BUILD_TESTING is ON.

include(FetchContent)

set(BUILD_GMOCK OFF CACHE BOOL "Disable gmock" FORCE)
set(INSTALL_GTEST OFF CACHE BOOL "Do not install gtest" FORCE)
set(gtest_force_shared_crt ON CACHE BOOL "Use shared CRT" FORCE)

FetchContent_Declare(
  googletest
  GIT_REPOSITORY https://github.com/google/googletest.git
  GIT_TAG v1.14.0
  GIT_SHALLOW TRUE
)

FetchContent_MakeAvailable(googletest)
