# FormulonWasm.cmake
#
# Configures the Emscripten WASM build targets.
#
# Two variants, switched via FM_WASM_VARIANT:
#
#   embind (default): the npm @libraz/formulon artifact. Compiles
#     src/wasm/parts/*.cpp + formulon_static, links with -lembind,
#     emits a MODULARIZE'd ES module factory (`createFormulon`). The
#     `parts/` split keeps the per-area binding surface manageable;
#     every TU is compiled with `-frtti` because embind keys its
#     type-registration tables on `typeid()`. FM_WASM_THREADS picks the
#     threading model:
#       OFF (default): no pthread, non-shared memory, so the module loads
#         without cross-origin isolation; `recalcParallel` degrades to
#         serial. Output: `<build>/formulon.{js,wasm}` (package root).
#       ON: -pthread / -sUSE_PTHREADS=1 / -sPTHREAD_POOL_SIZE=8 so the
#         parallel SCC scheduler runs on real Web Workers; needs
#         SharedArrayBuffer, hence COOP/COEP in a browser. Output:
#         `<build>/formulon_threads.{js,wasm}` (`@libraz/formulon/threads`).
#
#   capi: the formulon (PyPI) artifact, consumed by wasmtime-py. No
#     embind, no JS glue, no pthread. Compiles src/c_api/formulon_c.cpp
#     + formulon_static, link line is --no-entry + -sSTANDALONE_WASM=1,
#     exports the Python-required fm_* C ABI plus _malloc / _free. It
#     deliberately does not export `fm_workbook_recalc_parallel` or
#     `fm_workbook_recalc_full_parallel`: this no-pthread artifact
#     exposes only serial recalc to the Python wheel.
#     Output: `<build>/formulon_capi.wasm` (no companion .js).
#
# Activated only when FM_BUILD_WASM=ON, which itself is meaningful only
# under `emcmake`. The native build path (the default) MUST remain green
# without emscripten installed; nothing in this file should touch the
# build graph unless FM_BUILD_WASM is on AND EMSCRIPTEN is defined.
#
# Linker / codegen notes (embind variant)
# ---------------------------------------
#
# * Engine is built `-fno-exceptions -fno-rtti`. We keep that here
#   (`-sDISABLE_EXCEPTION_CATCHING=1`). Native methods return failures
#   through result envelopes; JS argument guards can throw before entry.
#   Caller-owned getters are caught in JS before returning to C++.
#
# * `MODULARIZE=1 + EXPORT_NAME=createFormulon + EXPORT_ES6=1` produces
#   an ES module factory: `import createFormulon from './formulon.js'`.
#   `ENVIRONMENT=web,worker,node` lets the same artifact run under all
#   three.
#
# * `FILESYSTEM=0` strips Emscripten's MEMFS shim; the engine never
#   touches files directly (callers hand in byte buffers via the C ABI).
#
# * `closure 0` because closure mangles embind's stub functions and
#   breaks the binding layer.
#
# * `INITIAL_MEMORY=33554432` (32 MiB) is the runtime heap, separate
#   from the `.wasm` code-size budget. `ALLOW_MEMORY_GROWTH=1` lets
#   large workbooks expand the heap, up to Emscripten's 2 GiB
#   `MAXIMUM_MEMORY` default (no larger value is set here: raising it
#   forces unsigned pointer arithmetic in the generated JS, a real
#   `.wasm`-size cost against the ceiling in CLAUDE.md's WASM size
#   policy, not something to opt into as a doc-comment correction).
#
# * `STACK_SIZE` is set explicitly on both variants rather than left to
#   Emscripten's default. See the shared block below for the measurement
#   the value comes from; it is not a round number picked for comfort.

if(NOT FM_BUILD_WASM)
  return()
endif()

if(NOT DEFINED EMSCRIPTEN)
  message(FATAL_ERROR
    "FM_BUILD_WASM=ON requires the Emscripten toolchain. "
    "Re-run cmake under `emcmake cmake ...` (see `make wasm`).")
endif()

set(FM_WASM_VARIANT "embind" CACHE STRING
  "WASM build variant: embind (npm) | capi (PyPI / wasmtime)")
set_property(CACHE FM_WASM_VARIANT PROPERTY STRINGS embind capi)

if(NOT FM_WASM_VARIANT STREQUAL "embind" AND NOT FM_WASM_VARIANT STREQUAL "capi")
  message(FATAL_ERROR
    "FM_WASM_VARIANT must be one of: embind, capi (got: ${FM_WASM_VARIANT})")
endif()

option(FM_WASM_THREADS "Link the embind variant with pthreads (@libraz/formulon/threads)" OFF)
if(FM_WASM_THREADS AND NOT FM_WASM_VARIANT STREQUAL "embind")
  message(FATAL_ERROR "FM_WASM_THREADS applies only to FM_WASM_VARIANT=embind")
endif()

# Output base name: artifact + JS glue land at <build>/<base>.{js,wasm}.
# Debug builds get a distinct base so debug + release can coexist when
# the build directory is reused.
if(CMAKE_BUILD_TYPE STREQUAL "Debug")
  set(_FM_WASM_BASE "formulon-debug")
else()
  set(_FM_WASM_BASE "formulon")
endif()

if(FM_WASM_VARIANT STREQUAL "capi")
  set(_FM_WASM_BASE "${_FM_WASM_BASE}_capi")
elseif(FM_WASM_THREADS)
  set(_FM_WASM_BASE "${_FM_WASM_BASE}_threads")
endif()

if(FM_WASM_VARIANT STREQUAL "embind")
  # embind glue is now sharded by API surface area under src/wasm/parts/.
  # Keep the list alphabetical so additions land in a predictable spot.
  add_executable(formulon_wasm
    src/wasm/parts/bindings_register.cpp
    src/wasm/parts/embind_common.cpp
    src/wasm/parts/workbook_auto_filter.cpp
    src/wasm/parts/workbook_callbacks.cpp
    src/wasm/parts/workbook_cells.cpp
    src/wasm/parts/workbook_cell_table_styles.cpp
    src/wasm/parts/workbook_cf.cpp
    src/wasm/parts/workbook_chart.cpp
    src/wasm/parts/workbook_chart_read.cpp
    src/wasm/parts/workbook_core.cpp
    src/wasm/parts/workbook_data_tables.cpp
    src/wasm/parts/workbook_drawing.cpp
    src/wasm/parts/workbook_shapes.cpp
    src/wasm/parts/workbook_slicers.cpp
    src/wasm/parts/workbook_misc.cpp
    src/wasm/parts/workbook_pivot.cpp
    src/wasm/parts/workbook_pivot_cache.cpp
    src/wasm/parts/workbook_reference.cpp
    src/wasm/parts/workbook_print.cpp
    src/wasm/parts/workbook_sheet_layout.cpp
    src/wasm/parts/workbook_styles.cpp
    src/wasm/parts/workbook_theme.cpp
    src/wasm/parts/workbook_ui.cpp
    # Generated by tools/codegen/gen_bindings.py from
    # tools/codegen/binding_manifest.yaml; see docs/refactor/04-phases-and-gates.md.
    src/wasm/generated/sheet_counts.cpp
    src/wasm/generated/styles_counts.cpp
    src/wasm/generated/workbook_counts.cpp
  )
  target_include_directories(formulon_wasm PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/src)
  target_link_libraries(formulon_wasm PRIVATE formulon_static)
else()
  # capi variant: drive the build from the C ABI translation unit. It's
  # already in formulon_core, but we need an executable entry so emcc
  # has something to link from. A trivial stub keeps the link graph
  # rooted; the actual fm_* exports come from the static archive.
  add_executable(formulon_wasm src/c_api/wasm_entry.cpp)
  target_include_directories(formulon_wasm PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/src)
  target_link_libraries(formulon_wasm PRIVATE formulon_static)
endif()

# The N-API registration table is the canonical declaration of primitive
# positional argument types. Generate the embind pre-js from that table so
# both JavaScript surfaces reject coercing values before embind enters a
# native method. The generated file is glue only and is therefore kept out
# of the wasm link itself.
if(FM_WASM_VARIANT STREQUAL "embind")
  find_package(Python3 COMPONENTS Interpreter REQUIRED)
  set(_FM_WASM_POSITIONAL_GUARD_SCRIPT
    "${CMAKE_CURRENT_SOURCE_DIR}/tools/codegen/gen_wasm_positional_guards.py")
  set(_FM_WASM_POSITIONAL_GUARD_PRE_JS
    "${CMAKE_CURRENT_BINARY_DIR}/generated/checked_arguments.pre.js")
  add_custom_command(
    OUTPUT "${_FM_WASM_POSITIONAL_GUARD_PRE_JS}"
    COMMAND "${Python3_EXECUTABLE}" "${_FM_WASM_POSITIONAL_GUARD_SCRIPT}"
            --node-source "${CMAKE_CURRENT_SOURCE_DIR}/src/node_addon/parts/workbook_class.cc"
            --wasm-source "${CMAKE_CURRENT_SOURCE_DIR}/src/wasm/parts/bindings_register.cpp"
            --output "${_FM_WASM_POSITIONAL_GUARD_PRE_JS}"
    DEPENDS
      "${_FM_WASM_POSITIONAL_GUARD_SCRIPT}"
      "${CMAKE_CURRENT_SOURCE_DIR}/src/node_addon/parts/workbook_class.cc"
      "${CMAKE_CURRENT_SOURCE_DIR}/src/wasm/parts/bindings_register.cpp"
      "${CMAKE_CURRENT_SOURCE_DIR}/src/node_addon/generated/facade_scope.inc"
      "${CMAKE_CURRENT_SOURCE_DIR}/src/c_api/generated/binding_param_names.inc"
    COMMENT "Generating WASM positional argument guards"
    VERBATIM
  )
  add_custom_target(formulon_wasm_positional_guards
    DEPENDS "${_FM_WASM_POSITIONAL_GUARD_PRE_JS}")
  add_dependencies(formulon_wasm formulon_wasm_positional_guards)
  set_property(TARGET formulon_wasm APPEND PROPERTY LINK_DEPENDS
    "${_FM_WASM_POSITIONAL_GUARD_PRE_JS}")
endif()

# Shared optimisation flags.
if(CMAKE_BUILD_TYPE STREQUAL "Debug")
  set(_FM_WASM_OPT_FLAGS "-O0;-g")
  set(_FM_WASM_GLUE_OPT_FLAGS "-O0;-g")
  set(_FM_WASM_LINK_OPT_FLAGS "-O0;-g;-sASSERTIONS=2")
else()
  # Full LTO grew the artifact in both link modes (-O3: raw +90 KB, Brotli +2 KB; -Oz: raw
  # +71 KB, Brotli -4 KB with the evaluator at -Oz) and took 3.5x the link time (measured
  # 2026-11-15); thinLTO grew it too, so LTO stays off.
  set(_FM_WASM_OPT_FLAGS "-O3")
  # The binding glue (formulon_wasm's own TUs) only marshals JS values, so it
  # is built -Oz while the hot core keeps -O3: raw -73 KB, Brotli -5 KB, bulk
  # cell I/O timings unchanged (measured 2026-10-21).
  set(_FM_WASM_GLUE_OPT_FLAGS "-Oz")
  # Core TUs, new ones included, are -Oz unless a hot pattern and no except pattern names them:
  # raw -61 KB, Brotli -19 KB (threads -14 KB); entry, recalc, load, save, row edits and
  # copySheet within noise (measured 2026-11-15). A pattern matching no source fails configure.
  set(_FM_WASM_HOT_PATTERNS
    "^src/eval/"
    "^src/parser/"
    "^src/utils/"
    "^src/cf/"
    "^src/io/cell_parser\\.cpp$"
    "^src/io/sheet_reader\\.cpp$"
    "^src/io/sst_reader\\.cpp$"
    "^src/io/styles_reader\\.cpp$"
    "^src/io/sax_xml_reader\\.cpp$"
    "^src/io/xml_utils\\.cpp$"
    "^src/io/xsd_double\\.cpp$"
    "^src/value\\.cpp$"
    "^src/sheet\\.cpp$"
    "^src/row_cells\\.cpp$"
    "^src/sheet_range_scan\\.cpp$"
    "^src/sheet_spill\\.cpp$"
    "^src/sheet_name\\.cpp$"
    "^src/sheet_name_casefold_data\\.cpp$"
    "^src/numeric_aggregate_kernels\\.cpp$"
    "^src/workbook\\.cpp$"
    "^src/workbook_formula_index\\.cpp$"
    "^src/workbook_cell_edit\\.cpp$"
    "^src/workbook_error_detail\\.cpp$"
    "^src/style_resolve\\.cpp$"
    "^src/style_region\\.cpp$"
    "^src/styles\\.cpp$"
    "^src/color_resolve\\.cpp$"
    "^src/excel_locale\\.cpp$"
    "^src/sbcs_codepage\\.cpp$"
    "^src/num_fmt_locale\\.cpp$"
    "^src/external_book\\.cpp$"
  )
  set(_FM_WASM_HOT_EXCEPT_PATTERNS
    "^src/parser/ast_dump\\.cpp$"
    "^src/parser/ast_format\\.cpp$"
    "^src/parser/ast_format_names\\.cpp$"
    "^src/parser/ast_format_parens\\.cpp$"
    "^src/parser/ast_format_r1c1\\.cpp$"
    "^src/parser/ast_shift\\.cpp$"
    "^src/parser/external_book_spelling\\.cpp$"
    "^src/parser/formula_prefix\\.cpp$"
    "^src/parser/literal_spelling\\.cpp$"
    "^src/parser/parse_error\\.cpp$"
    "^src/parser/ref_transforms\\.cpp$"
    "^src/parser/structured_ref_spelling\\.cpp$"
    "^src/cf/cf_authoring\\.cpp$"
    "^src/cf/cf_rule_traits\\.cpp$"
    "^src/cf/cf_x14_id\\.cpp$"
    "^src/eval/pivot_formula\\.cpp$"
    "^src/eval/formula_localize\\.cpp$"
    "^src/eval/ets_series\\.cpp$"
    "^src/eval/ets_detect\\.cpp$"
    "^src/eval/ets_calendar\\.cpp$"
    "^src/eval/ets_grid\\.cpp$"
    "^src/eval/date_text_parse\\.cpp$"
    "^src/eval/regex_lazy\\.cpp$"
    "^src/eval/regex_program\\.cpp$"
    "^src/eval/regex_replacement\\.cpp$"
    "^src/eval/locale_function_names\\.cpp$"
    "^src/eval/holt_winters\\.cpp$"
    "^src/eval/forecast_ets_lazy\\.cpp$"
    "^src/eval/financial_lazy\\.cpp$"
    "^src/eval/formula_entry_r1c1\\.cpp$"
    "^src/eval/coupon_schedule\\.cpp$"
    "^src/eval/database_lazy\\.cpp$"
    "^src/eval/getpivotdata_lazy\\.cpp$"
    "^src/eval/hypothesis_lazy\\.cpp$"
    "^src/eval/linest_lazy\\.cpp$"
    "^src/eval/regression_lazy\\.cpp$"
    "^src/eval/stats/"
    "^src/eval/builtins/financial"
    "^src/eval/builtins/engineering"
    "^src/eval/builtins/complex_num"
    "^src/eval/builtins/distributions"
    "^src/eval/builtins/cube"
    "^src/eval/builtins/web"
    "^src/eval/builtins/stats/stats_distributions"
    "^src/eval/builtins/math_combinatorics"
    "^src/eval/builtins/text_dbcs"
    "^src/eval/builtins/text_width"
    "^src/eval/text_format/"
    "^src/eval/groupby_pivotby/"
  )
  # The link runs Binaryen at -Oz: its -O3 inliner copied small functions into every caller, and the post-build
  # -Oz pass cannot undo that. With the third-party -Oz below: raw -97 KB, Brotli -14 KB, bulk workload within
  # noise (measured 2026-11-22).
  set(_FM_WASM_LINK_OPT_FLAGS "-Oz;-sASSERTIONS=0")
endif()

if(_FM_WASM_HOT_PATTERNS)
  foreach(_fm_pattern IN LISTS _FM_WASM_HOT_PATTERNS _FM_WASM_HOT_EXCEPT_PATTERNS)
    set(_fm_matched ${FORMULON_CORE_SOURCES})
    list(FILTER _fm_matched INCLUDE REGEX "${_fm_pattern}")
    if(NOT _fm_matched)
      message(FATAL_ERROR "WASM hot-list pattern '${_fm_pattern}' matches no core source")
    endif()
  endforeach()
  list(JOIN _FM_WASM_HOT_PATTERNS "|" _fm_hot_regex)
  list(JOIN _FM_WASM_HOT_EXCEPT_PATTERNS "|" _fm_except_regex)
  set(_FM_WASM_HOT_SOURCES ${FORMULON_CORE_SOURCES})
  list(FILTER _FM_WASM_HOT_SOURCES INCLUDE REGEX "${_fm_hot_regex}")
  list(FILTER _FM_WASM_HOT_SOURCES EXCLUDE REGEX "${_fm_except_regex}")
  set(_FM_WASM_COLD_SOURCES ${FORMULON_CORE_SOURCES})
  list(REMOVE_ITEM _FM_WASM_COLD_SOURCES ${_FM_WASM_HOT_SOURCES})
  list(LENGTH _FM_WASM_HOT_SOURCES _fm_hot_count)
  list(LENGTH FORMULON_CORE_SOURCES _fm_core_count)
  message(STATUS "WASM core TUs at -O3: ${_fm_hot_count} of ${_fm_core_count}; the rest -Oz")
  set_source_files_properties(${_FM_WASM_COLD_SOURCES} PROPERTIES COMPILE_OPTIONS "-Oz")
endif()

# Shadow-stack size, pinned rather than inherited.
#
# The engine's recursive formula walkers are bounded by the parser's nesting
# limits: Excel's own (`parser::kMaxFunctionNesting`,
# `kMaxGroupingParenDepth`, `kMaxOperatorNesting`) and the shared tree bound
# `parser::kMaxFormulaAstDepth`, which the Pratt parser counts its recursion
# against and `ast_format`, `ast_shift` and the XLSB `ptg_reader` re-check
# before walking a tree. Those limits are only a stack-safety guarantee if
# the stack they are sized against is known, and under Emscripten it is not
# known by default -- the default is 64 KiB, far inside what they permit.
#
# tests/wasm/stack_probe.cpp measures the deepest accepted shape of each kind
# through parse, evaluate, shift and format, and the runaway nesting the
# parser recurses into before refusing it, on every WASM build. The numbers
# live there rather than in this comment where they would rot.
#
# The pinned size is about 2.4x the probe's worst case. The margin matters more than
# usual here: the engine is built -fno-exceptions, and a shadow-stack
# overflow is not a C++-level failure in any case. Under the release link (`ASSERTIONS=0`, so
# `STACK_OVERFLOW_CHECK` is 0) the stack pointer simply walks past the limit
# into the data segment below it -- silent corruption for the first tens of
# KiB, then an untrappable out-of-bounds access. Neither outcome is
# reportable through `Expected`, so the size has to be right up front.
#
# Growing this does not grow the code section; it shifts the linear-memory
# layout, which moves the Brotli figure by a few hundred bytes.
set(_FM_WASM_STACK_SIZE 1572864)

if(FM_WASM_VARIANT STREQUAL "embind")
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
    "-sSTACK_SIZE=${_FM_WASM_STACK_SIZE}"
    "--closure=0"
  )
  list(APPEND _FM_WASM_COMMON_LINK_FLAGS
    "--pre-js=${_FM_WASM_POSITIONAL_GUARD_PRE_JS}")

  # Compile flags for the embind TU only. formulon_core is built
  # -fno-exceptions -fno-rtti (the project-wide stance). embind, however,
  # materially depends on RTTI: its type-registration tables key on
  # typeid() to map C++ types to JS wrappers. Building embind.cpp with
  # RTTI off but EMSCRIPTEN_HAS_UNBOUND_TYPE_NAMES=0 silences the
  # compile-time static_assert but leaves the runtime type-id strings
  # empty -- every bound entry then surfaces "Cannot call X due to
  # unbound types".
  #
  # We therefore enable RTTI *only* for the embind translation unit. The
  # formulon_static archive linked into this target was already built
  # with -fno-rtti; turning RTTI on for the binding glue does not
  # affect its codegen. Exceptions stay disabled both here and in the
  # linker step (-sDISABLE_EXCEPTION_CATCHING=1), and the binding routes
  # every failure through the JsStatus envelope rather than throwing.
  target_compile_options(formulon_wasm PRIVATE
    ${_FM_WASM_GLUE_OPT_FLAGS}
    -fno-exceptions
    -frtti
  )

  if(FM_WASM_THREADS)
    # -pthread is required at both compile and link time, and on
    # formulon_core too so the scheduler's atomics resolve to the
    # multi-threaded ABI. The pool is preallocated because recalcParallel
    # joins synchronously, which a browser main thread cannot do while a
    # worker is still being spawned. The 8-thread cap matches
    # kMaxAutoThreads in src/eval/scheduler.cpp.
    #
    # @size-budget-event: +~14 KB one-time pthread runtime.
    list(APPEND _FM_WASM_COMMON_LINK_FLAGS
      "-pthread"
      "-sUSE_PTHREADS=1"
      "-sPTHREAD_POOL_SIZE=8"
      # Worker stacks would follow STACK_SIZE implicitly; state it so the
      # coupling survives a toolchain whose default changes. This reserves
      # PTHREAD_POOL_SIZE x 1.5 MiB (12 MiB) of the runtime heap up front,
      # which is heap, not artifact size.
      "-sDEFAULT_PTHREAD_STACK_SIZE=${_FM_WASM_STACK_SIZE}"
      "-sBINARYEN_EXTRA_PASSES=--skip-pass=memory-packing"
    )
    # Shared memory initialises every data segment passively, so packing
    # them adds init code that compresses worse than the zeros it removes:
    # skipping it measured raw +19 KB / Brotli -2.3 KB (Brotli binds).
    set(_FM_WASM_OPT_SKIP_PASSES "--skip-pass=memory-packing")
    target_compile_options(formulon_wasm PRIVATE -pthread)
    target_compile_options(formulon_core PRIVATE -pthread)

    # formulon_core is compiled -pthread under this variant, so anything
    # linking the archive has to agree with it. The probe runs
    # single-threaded, hence a pool of zero.
    set(_FM_WASM_PROBE_THREAD_FLAGS -pthread)
    set(_FM_WASM_PROBE_THREAD_LINK_FLAGS "-pthread" "-sUSE_PTHREADS=1" "-sPTHREAD_POOL_SIZE=0")
  else()
    # Without pthreads `launch_thread` sees pthread_create refuse with
    # EAGAIN, and the scheduler evaluates the pass on the calling thread.
    set(_FM_WASM_PROBE_THREAD_FLAGS)
    set(_FM_WASM_PROBE_THREAD_LINK_FLAGS)
  endif()
else()
  # capi variant: standalone reactor-style WASM. No JS glue, no embind,
  # no pthread. The scheduler's std::thread calls compile but run
  # synchronously under emcc's no-pthread mode, so callers transparently
  # get serial recalc.
  #
  # Exports list: every fm_* function the Python binding may call, plus
  # _malloc / _free / _initialize so the host can allocate scratch
  # buffers for string arguments. The list is kept in a sibling file so
  # cmake doesn't have to be edited when the C ABI grows; the harness
  # in tools/wasm/capi_exports.txt is shared with the Python binding's
  # signature setup.
  file(STRINGS "${CMAKE_CURRENT_SOURCE_DIR}/tools/wasm/capi_exports.txt"
    _FM_WASM_CAPI_EXPORTS_RAW
    REGEX "^[^#].+")
  list(TRANSFORM _FM_WASM_CAPI_EXPORTS_RAW STRIP)
  list(FILTER _FM_WASM_CAPI_EXPORTS_RAW EXCLUDE REGEX "^$")
  list(TRANSFORM _FM_WASM_CAPI_EXPORTS_RAW PREPEND "_")
  list(JOIN _FM_WASM_CAPI_EXPORTS_RAW "," _FM_WASM_CAPI_EXPORTS)

  # Standalone WASM (reactor) link flags:
  # * NO_EXIT_RUNTIME / DYNAMIC_EXECUTION are JS-glue concepts and
  #   emcc rejects them under STANDALONE_WASM. Reactor builds (no
  #   main()) have EXIT_RUNTIME=false implicitly.
  # * ERROR_ON_UNDEFINED_SYMBOLS=0 because emcc still emits a small
  #   number of "private" runtime symbols (memory growth notifier,
  #   stack restore) that wasmtime resolves at instantiation via
  #   imports. The Python binding declares an env stub for these.
  # * WASM_BIGINT stays at its default (on): with it off, emcc splits the
  #   i64 params of WASI imports into i32 pairs, which wasmtime rejects.
  set(_FM_WASM_COMMON_LINK_FLAGS
    "--no-entry"
    "-sWASM=1"
    "-sSTANDALONE_WASM=1"
    "-sALLOW_MEMORY_GROWTH=1"
    "-sINITIAL_MEMORY=33554432"
    "-sFILESYSTEM=0"
    "-sDISABLE_EXCEPTION_CATCHING=1"
    "-sSUPPORT_LONGJMP=0"
    "-sMALLOC=emmalloc"
    "-sSTACK_SIZE=${_FM_WASM_STACK_SIZE}"
    "-sEXPORTED_FUNCTIONS=${_FM_WASM_CAPI_EXPORTS}"
    "--closure=0"
  )

  target_compile_options(formulon_wasm PRIVATE
    ${_FM_WASM_GLUE_OPT_FLAGS}
    -fno-exceptions
    -fno-rtti
  )

  set(_FM_WASM_PROBE_THREAD_FLAGS)
  set(_FM_WASM_PROBE_THREAD_LINK_FLAGS)
endif()

# Shadow-stack guard.
#
# tests/wasm/stack_probe.cpp measures what the parser's nesting limits
# cost the recursive walkers and fails if that has drifted past what the
# link reserves. It needs its own executable because peak stack usage is not
# observable through the binding surface, and it links with the shipped
# optimisation level and stack size so the numbers describe shipped codegen.
#
# `make wasm` and `make wasm-capi` build it; `make test-wasm` and
# `make python-test` run it. It is deliberately not part of
# `formulon_wasm`'s link: it is a second artifact next to the shipped
# one, and nothing stages it into a package. `make wasm-debug` does not build
# it either -- the budget it asserts is a statement about release codegen, and
# -O0 frames are several times wider.
add_executable(formulon_wasm_stack_probe tests/wasm/stack_probe.cpp)
target_include_directories(formulon_wasm_stack_probe PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/src)
target_link_libraries(formulon_wasm_stack_probe PRIVATE formulon_static)
target_compile_definitions(formulon_wasm_stack_probe PRIVATE
  FORMULON_WASM_EXPECTED_STACK_SIZE=${_FM_WASM_STACK_SIZE}
)
target_compile_options(formulon_wasm_stack_probe PRIVATE
  ${_FM_WASM_OPT_FLAGS}
  -fno-exceptions
  -fno-rtti
  ${_FM_WASM_PROBE_THREAD_FLAGS}
)
# No embind, no MODULARIZE: the probe has a main() and reports through its
# exit status. Everything that shapes codegen or the memory layout is taken
# from the shipped link.
target_link_options(formulon_wasm_stack_probe PRIVATE
  ${_FM_WASM_OPT_FLAGS}
  ${_FM_WASM_LINK_OPT_FLAGS}
  "-sSTACK_SIZE=${_FM_WASM_STACK_SIZE}"
  "-sINITIAL_MEMORY=33554432"
  "-sALLOW_MEMORY_GROWTH=1"
  "-sENVIRONMENT=node"
  "-sFILESYSTEM=0"
  "-sDISABLE_EXCEPTION_CATCHING=1"
  "-sSUPPORT_LONGJMP=0"
  "-sMALLOC=emmalloc"
  "-sEXIT_RUNTIME=1"
  ${_FM_WASM_PROBE_THREAD_LINK_FLAGS}
)
set_target_properties(formulon_wasm_stack_probe PROPERTIES
  OUTPUT_NAME "stack_probe"
  SUFFIX ".js"
)

# Define FORMULON_WASM for every TU under the WASM build. The
# read_ooxml SAX/DOM dispatch consults this macro: WASM builds get
# kSaxThresholdBytes = SIZE_MAX, which lets the linker
# dead-code-eliminate the streaming scanner (~15-20 KiB savings vs
# always-active SAX). Native builds keep the threshold at 256 KiB.
target_compile_definitions(formulon_core PRIVATE FORMULON_WASM=1)

# The standalone module asks its host for the local UTC offset, as WASI
# carries no time zone; the embind build keeps libc's localtime.
if(FM_WASM_VARIANT STREQUAL "capi")
  target_compile_definitions(formulon_core PRIVATE FORMULON_WASM_HOST_UTC_OFFSET=1)
  target_compile_definitions(formulon_wasm_stack_probe PRIVATE FORMULON_WASM_HOST_UTC_OFFSET=1)
endif()

# pugixml parses numbers with strtod and prints them with snprintf. Wrapping
# the symbols routes those calls to the engine's decimal parser
# (`__wrap_strtod` in src/utils/double_parse.cpp) and formatter
# (`__wrap_snprintf` in src/utils/number_text.cpp), which keeps libc's scanner,
# its printf machinery and the 128-bit float routines they need out of the
# module.
target_link_options(formulon_wasm PRIVATE
  ${_FM_WASM_OPT_FLAGS}
  ${_FM_WASM_LINK_OPT_FLAGS}
  ${_FM_WASM_COMMON_LINK_FLAGS}
  "-Wl,--wrap=strtod"
  "-Wl,--wrap=snprintf"
)

# Land the artifact at <build>/<base>.{js,wasm}. Emscripten infers the
# .wasm companion from the .js output name automatically. The capi
# variant uses --no-entry + STANDALONE_WASM so emscripten emits ONLY a
# .wasm (no companion .js); we still set SUFFIX so the cmake-tracked
# output path matches.
if(FM_WASM_VARIANT STREQUAL "embind")
  set_target_properties(formulon_wasm PROPERTIES
    OUTPUT_NAME "${_FM_WASM_BASE}"
    SUFFIX ".js"
  )
else()
  set_target_properties(formulon_wasm PROPERTIES
    OUTPUT_NAME "${_FM_WASM_BASE}"
    SUFFIX ".wasm"
  )
endif()

# Post-build: run wasm-opt with --converge so optimisation passes
# stabilise (Emscripten's link-time -Oz only invokes wasm-opt once).
# Located via the binaryen bin dir that ships with the active emcc
# toolchain. The pass is skipped under Debug to keep symbol info intact.
#
# The hint list has to cover two different toolchain layouts. An emsdk
# checkout puts emcc in `upstream/emscripten/` and wasm-opt one level up
# in `upstream/bin/`; a distro/homebrew emscripten keeps both in the same
# bin dir. `upstream/bin` is not on PATH after `emsdk_env.sh` either, so a
# missing hint is not rescued by the PATH fallback below -- it silently
# skipped the pass and shipped an unoptimised binary.
if(NOT CMAKE_BUILD_TYPE STREQUAL "Debug")
  get_filename_component(_FM_EMCC_DIR "${CMAKE_C_COMPILER}" DIRECTORY)
  find_program(WASM_OPT_EXECUTABLE
    NAMES wasm-opt
    HINTS "${_FM_EMCC_DIR}/../bin"
          "${_FM_EMCC_DIR}/../binaryen/bin"
          "${_FM_EMCC_DIR}/binaryen/bin"
    NO_DEFAULT_PATH
  )
  if(NOT WASM_OPT_EXECUTABLE)
    find_program(WASM_OPT_EXECUTABLE NAMES wasm-opt)
  endif()
  # Hard-fail rather than degrade quietly: the size ceilings are gated on
  # this artifact, so a build that skips the pass reports a size that is
  # not the one being judged.
  if(NOT WASM_OPT_EXECUTABLE)
    message(FATAL_ERROR
      "wasm-opt not found; the release WASM post-build pass cannot run and the "
      "artifact would miss the size the ceilings are measured against. It ships "
      "with the active toolchain -- pass -DWASM_OPT_EXECUTABLE=/path/to/wasm-opt "
      "if it lives outside the searched layouts.")
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
        ${_FM_WASM_OPT_SKIP_PASSES}
        $<TARGET_FILE_DIR:formulon_wasm>/${_FM_WASM_BASE}.wasm
        -o $<TARGET_FILE_DIR:formulon_wasm>/${_FM_WASM_BASE}.wasm
      VERBATIM
    )
  endif()
endif()

# Post-build: stamp the artifact with the emcc that produced it. The size
# ceilings are judged per compiler, so tools/bench/wasm_size_report.sh reads
# this sidecar rather than asking PATH -- that is what lets it flag a build
# made with a toolchain other than the pin in tools/wasm/emsdk-version.txt.
add_custom_command(TARGET formulon_wasm POST_BUILD
  COMMAND ${CMAKE_COMMAND}
    -DEMCC=${CMAKE_C_COMPILER}
    -DSTAMP_FILE=$<TARGET_FILE_DIR:formulon_wasm>/${_FM_WASM_BASE}.wasm.toolchain
    -P ${CMAKE_CURRENT_SOURCE_DIR}/tools/wasm/toolchain_stamp.cmake
  VERBATIM
)

# Post-build: print the artifact size (uncompressed + Brotli when
# available). This is informational only -- no enforcement here; the
# size gate lives in a separate bundle. Surface it on every successful
# wasm build so regressions are visible at a glance.
add_custom_command(TARGET formulon_wasm POST_BUILD
  COMMAND ${CMAKE_COMMAND}
    -DWASM_FILE=$<TARGET_FILE_DIR:formulon_wasm>/${_FM_WASM_BASE}.wasm
    -P ${CMAKE_CURRENT_SOURCE_DIR}/cmake/FormulonWasmSizeReport.cmake
  VERBATIM
)
