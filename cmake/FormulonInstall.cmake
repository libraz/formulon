# FormulonInstall.cmake
#
# Install rules for the shared C-ABI library: the library, its two public
# headers under include/formulon/, a CMake package (Formulon::formulon) and a
# pkg-config file. Only the shared library is installed; the third-party
# archives are linked into it privately. Active only when
# FM_BUILD_C_API_SHARED=ON.

if(NOT FM_BUILD_C_API_SHARED)
  return()
endif()

include(GNUInstallDirs)
include(CMakePackageConfigHelpers)

set_target_properties(formulon PROPERTIES
  EXPORT_NAME formulon
  VERSION ${formulon_VERSION}
  SOVERSION ${formulon_VERSION_MAJOR}.${formulon_VERSION_MINOR}
)

install(TARGETS formulon
  EXPORT FormulonTargets
  LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
  ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
  RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
  INCLUDES DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
)

install(FILES
  ${CMAKE_CURRENT_SOURCE_DIR}/src/c_api/formulon_c.h
  ${CMAKE_CURRENT_SOURCE_DIR}/src/c_api/formulon_status.h
  DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/formulon
)

set(FORMULON_INSTALL_CMAKEDIR ${CMAKE_INSTALL_LIBDIR}/cmake/Formulon)

install(EXPORT FormulonTargets
  NAMESPACE Formulon::
  DESTINATION ${FORMULON_INSTALL_CMAKEDIR}
)

configure_package_config_file(
  ${CMAKE_CURRENT_SOURCE_DIR}/cmake/FormulonConfig.cmake.in
  ${CMAKE_CURRENT_BINARY_DIR}/FormulonConfig.cmake
  INSTALL_DESTINATION ${FORMULON_INSTALL_CMAKEDIR}
)

# The C ABI may change at any minor release until 1.0.
write_basic_package_version_file(
  ${CMAKE_CURRENT_BINARY_DIR}/FormulonConfigVersion.cmake
  VERSION ${formulon_VERSION}
  COMPATIBILITY SameMinorVersion
)

install(FILES
  ${CMAKE_CURRENT_BINARY_DIR}/FormulonConfig.cmake
  ${CMAKE_CURRENT_BINARY_DIR}/FormulonConfigVersion.cmake
  DESTINATION ${FORMULON_INSTALL_CMAKEDIR}
)

# The prefix is resolved from the .pc file's own location so that an install
# relocated with `cmake --install --prefix` still works.
file(RELATIVE_PATH FORMULON_PC_PREFIX_REL
  "/${CMAKE_INSTALL_LIBDIR}/pkgconfig" "/")
set(FORMULON_PC_VERSION ${formulon_VERSION})
configure_file(
  ${CMAKE_CURRENT_SOURCE_DIR}/cmake/formulon.pc.in
  ${CMAKE_CURRENT_BINARY_DIR}/formulon.pc
  @ONLY
)
install(FILES ${CMAKE_CURRENT_BINARY_DIR}/formulon.pc
  DESTINATION ${CMAKE_INSTALL_LIBDIR}/pkgconfig
)
