# toolchain_stamp.cmake
#
# Record which Emscripten compiler produced a .wasm artifact, as a sidecar
# file next to it. Invoked as a POST_BUILD step from cmake/FormulonWasm.cmake:
#
#   cmake -DEMCC=<path> -DSTAMP_FILE=<artifact>.toolchain -P this-script
#
# tools/bench/wasm_size_report.sh reads the stamp instead of asking PATH, so
# the size report names the toolchain that actually built the artifact even
# when a different emcc is first on PATH by the time the report runs.
#
# A failure to query the compiler writes "unknown" rather than failing the
# build: the stamp is diagnostic, and losing it must not cost an artifact.

if(NOT DEFINED EMCC OR NOT DEFINED STAMP_FILE)
  message(FATAL_ERROR
    "toolchain_stamp.cmake requires -DEMCC=<path> and -DSTAMP_FILE=<path>")
endif()

execute_process(
  COMMAND "${EMCC}" --version
  OUTPUT_VARIABLE _stamp_output
  ERROR_VARIABLE _stamp_error
  RESULT_VARIABLE _stamp_result
  OUTPUT_STRIP_TRAILING_WHITESPACE
)

if(NOT _stamp_result EQUAL 0)
  message(WARNING
    "toolchain stamp: '${EMCC} --version' failed (${_stamp_result}): ${_stamp_error}")
  file(WRITE "${STAMP_FILE}" "unknown\n")
  return()
endif()

# emcc prints a multi-line banner; the first line carries the version.
string(REGEX MATCH "^[^\n]*" _stamp_line "${_stamp_output}")
file(WRITE "${STAMP_FILE}" "${_stamp_line}\n")
