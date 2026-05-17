# FormulonCommonDeps.cmake
#
# Provides formulon_apply_common_deps(target) which applies the canonical
# third-party dependency set shared by every artifact that wraps the engine
# core: the public `src/` include directory plus PUBLIC links to
# miniz / pugixml / pcre2 / double-conversion / Threads.
#
# The three current call sites are `formulon_core` (the OBJECT library),
# `formulon_static` (the in-tree STATIC archive that downstream consumers
# link), and the optional shared `formulon` C-ABI library. Each one needs
# the same usage requirements because $<TARGET_OBJECTS:...> does not
# propagate them and because consumers of the static archive expect the
# same transitive link set as the object library exports.

function(formulon_apply_common_deps target)
  if(NOT TARGET ${target})
    message(FATAL_ERROR
      "formulon_apply_common_deps: target '${target}' does not exist")
  endif()

  target_include_directories(${target}
    PUBLIC
      $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/src>
  )

  target_link_libraries(${target}
    PUBLIC
      formulon::miniz
      formulon::pugixml
      formulon::pcre2
      formulon::double_conversion
      Threads::Threads
  )
endfunction()
