# Out-of-tree consumer of the installed C ABI. Built twice from one source:
# once through the CMake package, once through pkg-config. Configure with
# CMAKE_PREFIX_PATH and PKG_CONFIG_PATH pointing at the install prefix.
cmake_minimum_required(VERSION 3.20)
project(formulon_install_consumer LANGUAGES C)

find_package(Formulon CONFIG REQUIRED)
add_executable(consumer_cmake main.c)
target_link_libraries(consumer_cmake PRIVATE Formulon::formulon)

find_package(PkgConfig REQUIRED)
pkg_check_modules(FORMULON_PC REQUIRED IMPORTED_TARGET formulon)
add_executable(consumer_pkgconfig main.c)
target_link_libraries(consumer_pkgconfig PRIVATE PkgConfig::FORMULON_PC)
set_target_properties(consumer_pkgconfig PROPERTIES
  BUILD_RPATH "${FORMULON_PC_LIBRARY_DIRS}")
