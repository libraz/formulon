# FormulonWasmSizeReport.cmake
#
# Standalone CMake script invoked from the `formulon_wasm` target's
# POST_BUILD step. Receives the path to the built `.wasm` file via the
# `WASM_FILE` variable and prints a one-line size summary including the
# Brotli-compressed size when `brotli` is available on PATH.
#
# Informational only: no enforcement here. The size gate is a separate
# bundle.

if(NOT WASM_FILE)
  message(STATUS "FormulonWasmSizeReport: WASM_FILE not provided; skipping")
  return()
endif()

if(NOT EXISTS "${WASM_FILE}")
  message(STATUS "FormulonWasmSizeReport: ${WASM_FILE} not found; skipping")
  return()
endif()

file(SIZE "${WASM_FILE}" _wasm_bytes)

# Format with one decimal of precision, in MiB.
math(EXPR _wasm_mib_int "${_wasm_bytes} / 1048576")
math(EXPR _wasm_mib_frac "(${_wasm_bytes} * 10 / 1048576) % 10")

# Locate brotli (best-effort).
find_program(BROTLI_EXECUTABLE brotli)
set(_brotli_msg "")
if(BROTLI_EXECUTABLE)
  set(_brotli_out "${WASM_FILE}.br")
  execute_process(
    COMMAND ${BROTLI_EXECUTABLE} --quality=11 --force --output=${_brotli_out} ${WASM_FILE}
    RESULT_VARIABLE _brotli_rc
    OUTPUT_QUIET
    ERROR_QUIET
  )
  if(_brotli_rc EQUAL 0 AND EXISTS "${_brotli_out}")
    file(SIZE "${_brotli_out}" _brotli_bytes)
    math(EXPR _brotli_kib "${_brotli_bytes} / 1024")
    set(_brotli_msg "; brotli=${_brotli_kib} KiB")
    file(REMOVE "${_brotli_out}")
  endif()
endif()

get_filename_component(_wasm_name "${WASM_FILE}" NAME)
message(STATUS "${_wasm_name}: ${_wasm_mib_int}.${_wasm_mib_frac} MiB (${_wasm_bytes} bytes)${_brotli_msg}")
