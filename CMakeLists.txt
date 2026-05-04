cmake_minimum_required(VERSION 3.20)
project(formulon VERSION 0.0.0 LANGUAGES CXX)

# ---------------------------------------------------------------------------
# Language standard and global settings
# ---------------------------------------------------------------------------
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)
set(CMAKE_POSITION_INDEPENDENT_CODE ON)
set(CMAKE_EXPORT_COMPILE_COMMANDS ON)

# Default to Debug when not explicitly specified.
if(NOT CMAKE_BUILD_TYPE AND NOT CMAKE_CONFIGURATION_TYPES)
  set(CMAKE_BUILD_TYPE Debug CACHE STRING "Build type" FORCE)
endif()

list(APPEND CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/cmake")

# Threads are required by the parallel recalc scheduler. `Threads::Threads`
# is INTERFACE-linked into `formulon_core` below; under Emscripten the WASM
# wiring in `cmake/FormulonWasm.cmake` adds `-pthread` / `-sUSE_PTHREADS`.
find_package(Threads REQUIRED)

# Reproducible build: honour SOURCE_DATE_EPOCH if present.
if(DEFINED ENV{SOURCE_DATE_EPOCH})
  set(FORMULON_SOURCE_DATE_EPOCH "$ENV{SOURCE_DATE_EPOCH}")
  message(STATUS "Formulon: SOURCE_DATE_EPOCH = ${FORMULON_SOURCE_DATE_EPOCH}")
endif()

# ---------------------------------------------------------------------------
# Build options (see backup/plans/06-build-packaging.md §6.1.2)
# ---------------------------------------------------------------------------
option(FM_BUILD_TESTING       "Build unit/property/oracle tests" ON)
option(FM_BUILD_WASM          "Build WASM target (requires emscripten)" OFF)
option(FM_BUILD_PYTHON_SHARED "Build shared lib for Python ctypes" OFF)
option(FM_BUILD_C_API_SHARED  "Build libformulon shared lib exporting the C ABI" OFF)
option(FM_BUILD_CLI           "Build formulon_cli binary" ON)
option(FM_BUILD_NODE_ADDON    "Build Node.js N-API addon (.node)" OFF)
option(FM_BUILD_FUZZ          "Build libFuzzer harnesses (clang only)" OFF)
option(FM_ENABLE_COVERAGE     "Enable gcov/lcov" OFF)
option(FM_USE_SYSTEM_DEPS     "Prefer system-installed third-party deps" OFF)
option(FM_ENABLE_PTHREAD      "Enable multithreaded recalc" ON)
# Tree-walker / VM parity harness: when ON, every `evaluate(...)` call also
# compiles the AST through the bytecode pipeline and asserts the VM result
# matches the tree-walker's. Defaults to OFF so production builds carry no
# overhead; the parity sweep test (`vm_parity_test.cpp`) flips it on.
option(FORMULON_VM_PARITY_CHECK "Enable tree-walker / VM parity assertions" OFF)

include(FormulonWarnings)

# ---------------------------------------------------------------------------
# Third-party dependencies (see backup/plans/06-build-packaging.md)
# ---------------------------------------------------------------------------
# miniz and pugixml are engine dependencies required by OOXML/XLSB I/O.
# pcre2 backs the REGEX* function family.
# When FM_USE_SYSTEM_DEPS=ON we expect system packages; otherwise we vendor
# via FetchContent (default).
if(FM_USE_SYSTEM_DEPS)
  find_package(miniz   CONFIG REQUIRED)
  find_package(pugixml CONFIG REQUIRED)
  find_package(pcre2   CONFIG REQUIRED COMPONENTS 8BIT)
  if(NOT TARGET formulon::miniz AND TARGET miniz::miniz)
    add_library(formulon::miniz ALIAS miniz::miniz)
  endif()
  if(NOT TARGET formulon::pugixml AND TARGET pugixml::pugixml)
    add_library(formulon::pugixml ALIAS pugixml::pugixml)
  endif()
  # PCRE2 system package exports `PCRE2::8BIT` (preferred) or the legacy
  # lowercase `pcre2-8` target name; alias whichever is present.
  if(NOT TARGET formulon::pcre2)
    if(TARGET PCRE2::8BIT)
      add_library(formulon::pcre2 ALIAS PCRE2::8BIT)
    elseif(TARGET pcre2-8-static)
      add_library(formulon::pcre2 ALIAS pcre2-8-static)
    elseif(TARGET pcre2-8)
      add_library(formulon::pcre2 ALIAS pcre2-8)
    else()
      message(FATAL_ERROR "FM_USE_SYSTEM_DEPS: pcre2 8-bit target not found")
    endif()
  endif()
else()
  include(cmake/FetchMiniz.cmake)
  include(cmake/FetchPugixml.cmake)
  include(cmake/FetchPcre2.cmake)
endif()

# ---------------------------------------------------------------------------
# Core object library (formulon_core)
# ---------------------------------------------------------------------------
# Core source list. Currently limited to the utilities, value, workbook,
# OOXML writer, and parser sources that have landed so far; new modules are
# appended here as they are implemented.
set(FORMULON_CORE_SOURCES
  src/utils/double_format.cpp
  src/utils/structured_log.cpp
  src/value.cpp
  src/sheet.cpp
  src/workbook.cpp
  src/io/cell_parser.cpp
  src/io/cf_reader.cpp
  src/io/cf_writer.cpp
  src/io/comments_reader.cpp
  src/io/comments_writer.cpp
  src/io/defined_names.cpp
  src/io/ooxml_reader.cpp
  src/io/ooxml_writer.cpp
  src/io/ooxml_writer_cell.cpp
  src/io/pivot_cache_reader.cpp
  src/io/pivot_cache_writer.cpp
  src/io/pivot_table_reader.cpp
  src/io/pivot_table_writer.cpp
  src/io/sax_xml_reader.cpp
  src/io/sheet_reader.cpp
  src/io/sst_reader.cpp
  src/io/styles_reader.cpp
  src/io/styles_writer.cpp
  src/io/tables_reader.cpp
  src/io/xml_escape.cpp
  src/io/zip_reader.cpp
  src/io/xlsb/cell_writer.cpp
  src/io/xlsb/func_id_table.cpp
  src/io/xlsb/ptg.cpp
  src/io/xlsb/reader.cpp
  src/io/xlsb/record.cpp
  src/io/xlsb/record_writer.cpp
  src/io/xlsb/sheet_writer.cpp
  src/io/xlsb/sst_writer.cpp
  src/io/xlsb/writer.cpp
  src/cf/cf_evaluator.cpp
  src/parser/tokenizer.cpp
  src/parser/ast.cpp
  src/parser/ast_dump.cpp
  src/parser/ast_format.cpp
  src/parser/ast_shift.cpp
  src/parser/parser.cpp
  src/parser/parser_atoms.cpp
  src/parser/parser_reference.cpp
  src/parser/ref_transforms.cpp
  src/pivot/pivot_evaluator.cpp
  src/pivot/pivot_index.cpp
  src/eval/coerce.cpp
  src/eval/compiler.cpp
  src/eval/optimizer.cpp
  src/eval/vm.cpp
  src/eval/utf8_length.cpp
  src/eval/text_ops.cpp
  src/eval/rng.cpp
  src/eval/dep_graph.cpp
  src/eval/dep_extractor.cpp
  src/eval/dirty_set.cpp
  src/eval/iterative_solver.cpp
  src/eval/recalc_engine.cpp
  src/eval/scheduler.cpp
  src/eval/volatile_tracker.cpp
  src/eval/function_registry.cpp
  src/eval/special_forms_catalog.cpp
  src/eval/builtins.cpp
  src/eval/builtins/aggregate.cpp
  src/eval/builtins/complex_num.cpp
  src/eval/builtins/cube.cpp
  src/eval/builtins/datetime.cpp
  src/eval/builtins/distributions.cpp
  src/eval/builtins/dynamic_array.cpp
  src/eval/builtins/engineering.cpp
  src/eval/builtins/engineering_convert.cpp
  src/eval/builtins/engineering_special.cpp
  src/eval/builtins/financial.cpp
  src/eval/builtins/financial_accrual.cpp
  src/eval/builtins/financial_bond_simple.cpp
  src/eval/builtins/financial_coupon.cpp
  src/eval/builtins/financial_depreciation.cpp
  src/eval/builtins/financial_duration.cpp
  src/eval/builtins/financial_misc.cpp
  src/eval/builtins/financial_oddfprice.cpp
  src/eval/builtins/financial_oddfyield.cpp
  src/eval/builtins/financial_oddlprice.cpp
  src/eval/builtins/financial_oddlyield.cpp
  src/eval/builtins/financial_price.cpp
  src/eval/builtins/financial_rates.cpp
  src/eval/builtins/financial_yield.cpp
  src/eval/builtins/info.cpp
  src/eval/builtins/logical.cpp
  src/eval/builtins/math.cpp
  src/eval/builtins/math_combinatorics.cpp
  src/eval/builtins/math_rng.cpp
  src/eval/builtins/math_trig.cpp
  src/eval/builtins/references.cpp
  src/eval/builtins/service_stubs.cpp
  src/eval/builtins/stats.cpp
  src/eval/builtins/stats/stats_distributions.cpp
  src/eval/builtins/stats/stats_distributions_misc.cpp
  src/eval/builtins/subtotal.cpp
  src/eval/builtins/text.cpp
  src/eval/builtins/text_bahttext.cpp
  src/eval/builtins/text_dbcs.cpp
  src/eval/builtins/text_detail.cpp
  src/eval/builtins/text_modern.cpp
  src/eval/builtins/text_format.cpp
  src/eval/builtins/text_width.cpp
  src/eval/builtins/web.cpp
  src/eval/text_format/number_format.cpp
  src/eval/text_format/number_format_tokenizer.cpp
  src/eval/text_format/number_format_render.cpp
  src/eval/stats/special_functions.cpp
  src/eval/tree_walker.cpp
  src/eval/eval_context.cpp
  src/eval/eval_state.cpp
  src/eval/criteria.cpp
  src/eval/jp_fold.cpp
  src/eval/jis0208_table.cpp
  src/eval/jis0208_table_lookup.cpp
  src/eval/coupon_schedule.cpp
  src/eval/date_time.cpp
  src/eval/date_text_parse.cpp
  src/eval/range_args.cpp
  src/eval/scalar_ops.cpp
  src/eval/special_forms_lazy.cpp
  src/eval/structured_ref.cpp
  src/eval/aggregate_lazy.cpp
  src/eval/areas_lazy.cpp
  src/eval/cell_lazy.cpp
  src/eval/conditional_aggregates.cpp
  src/eval/database_lazy.cpp
  src/eval/dynamic_array_lazy.cpp
  src/eval/groupby_pivotby_lazy.cpp
  src/eval/lambda_format.cpp
  src/eval/lambda_helpers_lazy.cpp
  src/eval/matrix_ops_lazy.cpp
  src/eval/shape_ops_lazy.cpp
  src/eval/workdays_lazy.cpp
  src/eval/financial_lazy.cpp
  src/eval/forecast_ets_lazy.cpp
  src/eval/getpivotdata_lazy.cpp
  src/eval/linest_lazy.cpp
  src/eval/textsplit_lazy.cpp
  src/eval/trimrange_lazy.cpp
  src/eval/reference_lazy.cpp
  src/eval/regression_lazy.cpp
  src/eval/series_sum_lazy.cpp
  src/eval/rank_lazy.cpp
  src/eval/regex_lazy.cpp
  src/eval/hypothesis_lazy.cpp
  src/eval/info_lazy.cpp
  src/eval/phonetic_lazy.cpp
  src/eval/lookups/classic.cpp
  src/eval/lookups/xlookup.cpp
  src/c_api/formulon_c.cpp
)

add_library(formulon_core OBJECT ${FORMULON_CORE_SOURCES})

target_include_directories(formulon_core
  PUBLIC
    $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/src>
)

# Engine core depends on miniz (zip/deflate), pugixml (XML), and pcre2
# (REGEX* function family). Link PUBLIC so downstream targets
# (formulon_static, tests, CLI) inherit include dirs and transitive link
# requirements.
target_link_libraries(formulon_core
  PUBLIC
    formulon::miniz
    formulon::pugixml
    formulon::pcre2
    Threads::Threads
)

formulon_apply_warnings(formulon_core)

if(FORMULON_VM_PARITY_CHECK)
  target_compile_definitions(formulon_core PUBLIC FORMULON_VM_PARITY_CHECK=1)
endif()

# ---------------------------------------------------------------------------
# Static library aggregating the core objects
# ---------------------------------------------------------------------------
add_library(formulon_static STATIC $<TARGET_OBJECTS:formulon_core>)

target_include_directories(formulon_static
  PUBLIC
    $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/src>
)

# $<TARGET_OBJECTS:...> does not propagate usage requirements, so replicate
# the third-party link dependencies here for downstream consumers.
target_link_libraries(formulon_static
  PUBLIC
    formulon::miniz
    formulon::pugixml
    formulon::pcre2
    Threads::Threads
)

# No per-target warning flags: formulon_static only aggregates objects.

# ---------------------------------------------------------------------------
# Optional shared library exporting the C ABI (libformulon).
#
# Default OFF: in-tree consumers (CLI, WASM bindings, tests) link the static
# archive directly. The shared variant exists for out-of-process language
# bindings (Python ctypes, future REST host, etc.) that need to dlopen a
# single artifact and resolve only the `fm_*` symbols. The C API
# translation unit is built with default visibility for its `FM_API`
# entry points; everything else stays hidden because the rest of the
# core is compiled with `-fvisibility=hidden`.
# ---------------------------------------------------------------------------
if(FM_BUILD_C_API_SHARED)
  add_library(formulon SHARED $<TARGET_OBJECTS:formulon_core>)

  target_include_directories(formulon
    PUBLIC
      $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/src>
  )

  target_link_libraries(formulon
    PUBLIC
      formulon::miniz
      formulon::pugixml
      formulon::pcre2
      Threads::Threads
  )
endif()

# ---------------------------------------------------------------------------
# CLI target
#
# `formulon_cli` is a thin argv dispatcher around the stable C ABI
# declared in `src/c_api/formulon_c.h`. It links to `formulon_static`
# and routes each subcommand into its own TU under `src/cli/`. The
# binary is intentionally kept dependency-light: the only entry it
# touches outside the C ABI is `cli/render` (locale-independent value
# rendering), which lives next to the command sources.
# ---------------------------------------------------------------------------
if(FM_BUILD_CLI)
  set(FORMULON_CLI_SOURCES
    src/cli/main.cpp
    src/cli/eval_cmd.cpp
    src/cli/recalc_cmd.cpp
    src/cli/dump_cmd.cpp
    src/cli/render.cpp
  )
  add_executable(formulon_cli ${FORMULON_CLI_SOURCES})
  target_link_libraries(formulon_cli PRIVATE formulon_static)
  target_include_directories(formulon_cli PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/src)
  set_target_properties(formulon_cli PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin"
  )
  formulon_apply_warnings(formulon_cli)
endif()

# ---------------------------------------------------------------------------
# WASM target (Emscripten / embind bindings)
# ---------------------------------------------------------------------------
# The actual target wiring lives in cmake/FormulonWasm.cmake so the top
# of the build remains untouched on native CI (where emscripten is not
# installed). The script is a no-op unless FM_BUILD_WASM=ON.
include(cmake/FormulonWasm.cmake)

# ---------------------------------------------------------------------------
# Node.js N-API addon target
# ---------------------------------------------------------------------------
# The actual target wiring lives in cmake/FormulonNodeAddon.cmake; the
# script is a no-op unless FM_BUILD_NODE_ADDON=ON. Linking happens
# against `formulon_static`, so multithreaded recalc is inherited
# automatically from the engine.
if(FM_BUILD_NODE_ADDON)
  include(cmake/FormulonNodeAddon.cmake)
endif()

# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------
if(FM_BUILD_TESTING)
  enable_testing()
  add_subdirectory(tests)
endif()
