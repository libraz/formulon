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
option(FM_BUILD_CLI           "Build formulon_cli binary" ON)
option(FM_BUILD_FUZZ          "Build libFuzzer harnesses (clang only)" OFF)
option(FM_ENABLE_COVERAGE     "Enable gcov/lcov" OFF)
option(FM_USE_SYSTEM_DEPS     "Prefer system-installed third-party deps" OFF)
option(FM_ENABLE_PTHREAD      "Enable multithreaded recalc" ON)

include(FormulonWarnings)

# ---------------------------------------------------------------------------
# Third-party dependencies (see backup/plans/06-build-packaging.md)
# ---------------------------------------------------------------------------
# miniz and pugixml are engine dependencies required by OOXML/XLSB I/O.
# When FM_USE_SYSTEM_DEPS=ON we expect system packages; otherwise we vendor
# via FetchContent (default).
if(FM_USE_SYSTEM_DEPS)
  find_package(miniz   CONFIG REQUIRED)
  find_package(pugixml CONFIG REQUIRED)
  if(NOT TARGET formulon::miniz AND TARGET miniz::miniz)
    add_library(formulon::miniz ALIAS miniz::miniz)
  endif()
  if(NOT TARGET formulon::pugixml AND TARGET pugixml::pugixml)
    add_library(formulon::pugixml ALIAS pugixml::pugixml)
  endif()
else()
  include(cmake/FetchMiniz.cmake)
  include(cmake/FetchPugixml.cmake)
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
  src/io/ooxml_writer.cpp
  src/parser/tokenizer.cpp
  src/parser/ast.cpp
  src/parser/ast_dump.cpp
  src/parser/parser.cpp
  src/parser/parser_atoms.cpp
  src/parser/parser_reference.cpp
  src/eval/coerce.cpp
  src/eval/utf8_length.cpp
  src/eval/text_ops.cpp
  src/eval/function_registry.cpp
  src/eval/special_forms_catalog.cpp
  src/eval/builtins.cpp
  src/eval/builtins/aggregate.cpp
  src/eval/builtins/complex_num.cpp
  src/eval/builtins/cube.cpp
  src/eval/builtins/datetime.cpp
  src/eval/builtins/distributions.cpp
  src/eval/builtins/engineering.cpp
  src/eval/builtins/engineering_convert.cpp
  src/eval/builtins/engineering_special.cpp
  src/eval/builtins/financial.cpp
  src/eval/builtins/financial_accrual.cpp
  src/eval/builtins/financial_bond_simple.cpp
  src/eval/builtins/financial_coupon.cpp
  src/eval/builtins/financial_depreciation.cpp
  src/eval/builtins/financial_misc.cpp
  src/eval/builtins/financial_rates.cpp
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
  src/eval/builtins/text_detail.cpp
  src/eval/builtins/text_dbcs.cpp
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
  src/eval/coupon_schedule.cpp
  src/eval/date_time.cpp
  src/eval/date_text_parse.cpp
  src/eval/range_args.cpp
  src/eval/special_forms_lazy.cpp
  src/eval/areas_lazy.cpp
  src/eval/conditional_aggregates.cpp
  src/eval/database_lazy.cpp
  src/eval/shape_ops_lazy.cpp
  src/eval/workdays_lazy.cpp
  src/eval/financial_lazy.cpp
  src/eval/reference_lazy.cpp
  src/eval/regression_lazy.cpp
  src/eval/series_sum_lazy.cpp
  src/eval/rank_lazy.cpp
  src/eval/hypothesis_lazy.cpp
  src/eval/info_lazy.cpp
  src/eval/lookups/classic.cpp
  src/eval/lookups/xlookup.cpp
)

add_library(formulon_core OBJECT ${FORMULON_CORE_SOURCES})

target_include_directories(formulon_core
  PUBLIC
    $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/src>
)

# Engine core depends on miniz (zip/deflate) and pugixml (XML). Link PUBLIC
# so downstream targets (formulon_static, tests, CLI) inherit include dirs
# and transitive link requirements.
target_link_libraries(formulon_core
  PUBLIC
    formulon::miniz
    formulon::pugixml
)

formulon_apply_warnings(formulon_core)

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
)

# No per-target warning flags: formulon_static only aggregates objects.

# ---------------------------------------------------------------------------
# CLI target (option kept ON by default, but the binary is not built until
# tools/cli lands. Placeholder branch to stay honest with the option
# contract.)
# ---------------------------------------------------------------------------
if(FM_BUILD_CLI)
  message(STATUS "Formulon: FM_BUILD_CLI=ON (CLI target not yet wired; placeholder)")
endif()

# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------
if(FM_BUILD_TESTING)
  enable_testing()
  add_subdirectory(tests)
endif()
