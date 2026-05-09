# FormulonWasm.cmake
#
# Configures the Emscripten / embind build target (`formulon_wasm`).
#
# Activated only when `FM_BUILD_WASM=ON`, which is itself meaningful only
# under `emcmake`. The native build path (the default) MUST remain green
# without emscripten installed; nothing in this file should touch the
# build graph unless `FM_BUILD_WASM` is on AND `EMSCRIPTEN` is defined.
#
# Output layout (relative to the build dir): `formulon.js` + `formulon.wasm`
# for Release, `formulon-debug.js` + `formulon-debug.wasm` for Debug.
#
# Linker / codegen notes
# ----------------------
#
# * The whole engine is built `-fno-exceptions -fno-rtti`. We keep that
#   here (`-sDISABLE_EXCEPTION_CATCHING=1`); the embind binding never
#   throws because `JsWorkbook` and friends route every failure through
#   a `{ ok, status, message }` envelope.
#
# * `MODULARIZE=1 + EXPORT_NAME=createFormulon + EXPORT_ES6=1` produces
#   an ES module factory: `import createFormulon from './formulon.js'`.
#   Setting `ENVIRONMENT=web,worker,node` lets the same artifact run
#   under all three.
#
# * `FILESYSTEM=0` strips Emscripten's MEMFS shim; the engine never
#   touches files directly (callers hand in byte buffers via the C ABI).
#
# * `closure 0` because closure mangles embind's stub functions and
#   breaks the binding layer. Bundle 5+ size optimisation may revisit.
#
# * `INITIAL_MEMORY=33554432` (32 MiB) is the runtime heap, separate
#   from the `.wasm` code-size budget. `ALLOW_MEMORY_GROWTH=1`
#   lets large workbooks expand the heap up to the 4 GiB limit.

if(NOT FM_BUILD_WASM)
  return()
endif()

if(NOT DEFINED EMSCRIPTEN)
  message(FATAL_ERROR
    "FM_BUILD_WASM=ON requires the Emscripten toolchain. "
    "Re-run cmake under `emcmake cmake ...` (see `make wasm`).")
endif()

# Output base name: artifact + JS glue land at <build>/<base>.{js,wasm}.
# Debug builds get a distinct base so debug + release can coexist when
# the build directory is reused.
if(CMAKE_BUILD_TYPE STREQUAL "Debug")
  set(_FM_WASM_BASE "formulon-debug")
else()
  set(_FM_WASM_BASE "formulon")
endif()

add_executable(formulon_wasm src/wasm/embind.cpp)
target_include_directories(formulon_wasm PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/src)
target_link_libraries(formulon_wasm PRIVATE formulon_static)

# embind requires `-lembind` at link time. Compile flags mirror the
# project-wide stance (`-fno-exceptions -fno-rtti`) plus optimisation
# tier; link flags configure the JS / module surface.
#
# Threads (parallel recalc scheduler): `-pthread` is required at both
# compile and link time under Emscripten. The link side additionally
# wants `-sUSE_PTHREADS=1` and a non-zero `PTHREAD_POOL_SIZE` so the
# runtime preallocates worker Web Workers up front. The 8-thread cap
# matches `kMaxAutoThreads` in `src/eval/scheduler.cpp`.
#
# @size-budget-event: +~14 KB one-time pthread runtime (approved by
# wasm-size-guardian). Future scheduler additions will not re-add this
# cost.
set(_FM_WASM_COMMON_LINK_FLAGS
  "-lembind"
  "-sWASM=1"
  "-sMODULARIZE=1"
  "-sEXPORT_NAME=createFormulon"
  "-sEXPORT_ES6=1"
  "-sENVIRONMENT=web,worker,node"
  "-sALLOW_MEMORY_GROWTH=1"
  "-sINITIAL_MEMORY=33554432"
  "-sNO_EXIT_RUNTIME=1"
  "-sFILESYSTEM=0"
  "-sDISABLE_EXCEPTION_CATCHING=1"
  "-sSUPPORT_LONGJMP=0"
  "-sDYNAMIC_EXECUTION=0"
  "-sWASM_BIGINT=0"
  "-sMALLOC=emmalloc"
  "-pthread"
  "-sUSE_PTHREADS=1"
  "-sPTHREAD_POOL_SIZE=8"
  "--closure=0"
)

if(CMAKE_BUILD_TYPE STREQUAL "Debug")
  set(_FM_WASM_OPT_FLAGS "-O0;-g")
  set(_FM_WASM_LINK_OPT_FLAGS "-O0;-g;-sASSERTIONS=2")
else()
  # -O3 outperforms -Oz on this codebase (measured 2026-05): -Oz with
  # thinLTO grew the artifact by ~270 KiB because the engine's hot lazy
  # dispatch tables benefit materially from -O3's inlining heuristics.
  # The size win comes instead from the wasm-opt post-link pass below.
  #
  # Full LTO (`-flto=full`) was retried 2026-05-09 with `-O3` + post-link
  # `wasm-opt -Oz` to address the 2026-05 thinLTO regression. Empirical
  # delta: +345 KiB (1.75 MiB -> 2.08 MiB). Same root cause: cross-TU
  # inlining over the lazy dispatch tables expands more code than ICF
  # reclaims. Both LTO modes are therefore confirmed off-limits for this
  # codebase under current heuristics.
  set(_FM_WASM_OPT_FLAGS "-O3")
  set(_FM_WASM_LINK_OPT_FLAGS "-O3;-sASSERTIONS=0")
endif()

# Compile flags for the embind TU only.
#
# `formulon_core` is built `-fno-exceptions -fno-rtti` (the project-wide
# stance). embind, however, materially depends on RTTI: its
# type-registration tables key on `typeid()` to map C++ types to JS
# wrappers. Building embind.cpp with RTTI off but
# EMSCRIPTEN_HAS_UNBOUND_TYPE_NAMES=0 silences the compile-time
# static_assert but leaves the runtime type-id strings empty — every
# bound entry then surfaces "Cannot call X due to unbound types".
#
# We therefore enable RTTI *only* for the embind translation unit. The
# `formulon_static` archive linked into this target was already built
# with `-fno-rtti`; turning RTTI on for the binding glue does not
# affect its codegen. Exceptions stay disabled both here and in the
# linker step (`-sDISABLE_EXCEPTION_CATCHING=1`), and the binding
# routes every failure through the `JsStatus` envelope rather than
# throwing.
target_compile_options(formulon_wasm PRIVATE
  ${_FM_WASM_OPT_FLAGS}
  -fno-exceptions
  -frtti
  -pthread
)

# Threading is wired in at the workbook scheduler. Under Emscripten the
# core library MUST also be built with `-pthread` so atomics in the
# scheduler's worker pool resolve to the multi-threaded ABI. Native
# builds do not need a special flag — `Threads::Threads` propagates the
# host pthread requirements.
target_compile_options(formulon_core PRIVATE -pthread)

# Define `FORMULON_WASM` for every TU under the WASM build. The
# `read_ooxml` SAX/DOM dispatch consults this macro: WASM builds get
# `kSaxThresholdBytes = SIZE_MAX`, which lets the linker
# dead-code-eliminate the streaming scanner (~15-20 KiB savings vs
# always-active SAX). Native builds keep the threshold at 256 KiB.
target_compile_definitions(formulon_core PRIVATE FORMULON_WASM=1)

target_link_options(formulon_wasm PRIVATE
  ${_FM_WASM_OPT_FLAGS}
  ${_FM_WASM_LINK_OPT_FLAGS}
  ${_FM_WASM_COMMON_LINK_FLAGS}
)

# Land the artifact at <build>/<base>.{js,wasm}. Emscripten infers the
# `.wasm` companion from the `.js` output name automatically.
set_target_properties(formulon_wasm PROPERTIES
  OUTPUT_NAME "${_FM_WASM_BASE}"
  SUFFIX ".js"
)

# Post-build: run wasm-opt with --converge so optimisation passes
# stabilise (Emscripten's link-time -Oz only invokes wasm-opt once).
# Located via the binaryen bin dir that ships with the active emcc
# toolchain. The pass is skipped under Debug to keep symbol info intact.
if(NOT CMAKE_BUILD_TYPE STREQUAL "Debug")
  get_filename_component(_FM_EMCC_DIR "${CMAKE_C_COMPILER}" DIRECTORY)
  find_program(WASM_OPT_EXECUTABLE
    NAMES wasm-opt
    HINTS "${_FM_EMCC_DIR}/../binaryen/bin"
          "${_FM_EMCC_DIR}/binaryen/bin"
    NO_DEFAULT_PATH
  )
  if(NOT WASM_OPT_EXECUTABLE)
    find_program(WASM_OPT_EXECUTABLE NAMES wasm-opt)
  endif()
  if(WASM_OPT_EXECUTABLE)
    add_custom_command(TARGET formulon_wasm POST_BUILD
      COMMAND ${WASM_OPT_EXECUTABLE}
        -Oz --converge
        --strip-debug --strip-producers
        --vacuum --dce --rse
        --enable-bulk-memory --enable-sign-ext --enable-mutable-globals
        --enable-threads --enable-nontrapping-float-to-int --enable-multivalue
        --enable-reference-types
        $<TARGET_FILE_DIR:formulon_wasm>/${_FM_WASM_BASE}.wasm
        -o $<TARGET_FILE_DIR:formulon_wasm>/${_FM_WASM_BASE}.wasm
      VERBATIM
    )
  endif()
endif()

# Post-build: print the artifact size (uncompressed + Brotli when
# available). This is informational only — no enforcement here; the
# size gate lives in a separate bundle. Surface it on every successful
# wasm build so regressions are visible at a glance.
add_custom_command(TARGET formulon_wasm POST_BUILD
  COMMAND ${CMAKE_COMMAND}
    -DWASM_FILE=$<TARGET_FILE_DIR:formulon_wasm>/${_FM_WASM_BASE}.wasm
    -P ${CMAKE_CURRENT_SOURCE_DIR}/cmake/FormulonWasmSizeReport.cmake
  VERBATIM
)
