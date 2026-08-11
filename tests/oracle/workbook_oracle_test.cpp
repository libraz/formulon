//
// Parameterized gtest skeleton for the workbook oracle track.
//
// The workbook oracle covers workbook-level features that are NOT formula
// results -- pivot tables and print areas. Each parameter is a
// (suite, case_id) pair loaded via `load_workbook_oracle_cases` from
// `tests/oracle/golden_wb/*.golden.json`. The build wires the directory
// through the compile-time define `FORMULON_WORKBOOK_ORACLE_GOLDEN_DIR`.
// When the directory is empty (the expected state before the workbook
// oracle generator has been run on a Windows host) the parameter vector
// is empty and gtest registers the suite with zero instantiations, so the
// build stays green.
//
// Both verifiers are implemented: a pivot case is rebuilt, evaluated and
// laid out, then its rendered grid is diffed against `expect.pivot.grid`;
// a print case is rebuilt and paginated, then the `PaginationResult` is
// diffed against `expect.print`. With `golden_wb/` empty the parameter
// vector is empty, so the suite registers zero cases and the build stays
// green; feature-less cases are skipped.

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <iterator>
#include <map>
#include <string>
#include <utility>
#include <vector>

#include "eval/pivot_locale.h"
#include "gtest/gtest.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "print/pagination.h"
#include "print/print_area.h"
#include "tests/oracle/json_reader.h"
#include "tests/oracle/workbook_builder.h"
#include "tests/oracle/workbook_oracle_runner.h"
#include "value.h"

#ifndef FORMULON_WORKBOOK_ORACLE_GOLDEN_DIR
#define FORMULON_WORKBOOK_ORACLE_GOLDEN_DIR ""
#endif

namespace formulon {
namespace tests {
namespace oracle {
namespace {

// ---------------------------------------------------------------------------
// Parameter provider
// ---------------------------------------------------------------------------

const std::vector<WorkbookOracleCase>& workbook_oracle_cases() {
  // Loaded once at first call. An empty / absent golden_wb directory
  // yields an empty vector, so the parameterized suite registers zero
  // cases and the build stays green.
  //
  // Variant goldens (tests/oracle/variants/<tag>/golden_wb/) are appended
  // after the primary set. Each variant case inherits its tag from the
  // load call; the parameter-name printer suffixes `__<tag>` so primary
  // and variant entries never collide. The primary build configures no
  // variant dirs, so the appended sequence is empty and the parameter
  // list is identical to the primary-only flow; the variant binary
  // passes an empty primary dir and loads ONLY the variant trees.
  static const std::vector<WorkbookOracleCase> cached = []() {
    std::vector<WorkbookOracleCase> all = load_workbook_oracle_cases(FORMULON_WORKBOOK_ORACLE_GOLDEN_DIR, "");
    for (const auto& [tag, dir] : configured_workbook_variant_dirs()) {
      auto vc = load_workbook_oracle_cases(dir, tag);
      all.insert(all.end(), std::make_move_iterator(vc.begin()), std::make_move_iterator(vc.end()));
    }
    return all;
  }();
  return cached;
}

// ---------------------------------------------------------------------------
// Pivot grid comparison
// ---------------------------------------------------------------------------

// One rendered pivot cell, addressed relative to the pivot anchor.
struct GridCell {
  std::uint32_t r = 0;
  std::uint32_t c = 0;
  // Comparison is performed on the stringified value so number / text /
  // bool / blank / error cells share one normalised representation. Excel
  // goldens and the engine both render through the same display rules, so a
  // string match is the appropriate equality here.
  std::string value;
};

// Renders a `Value` into the canonical string the golden grid stores.
std::string render_value(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Blank:
      return "";
    case ValueKind::Bool:
      return v.as_boolean() ? "TRUE" : "FALSE";
    case ValueKind::Text:
      return std::string(v.as_text());
    case ValueKind::Number: {
      // Integers render without a trailing ".0"; other doubles use the
      // shortest %g form, matching how the golden generator normalises
      // numeric pivot cells.
      const double n = v.as_number();
      if (n == static_cast<double>(static_cast<long long>(n))) {
        return std::to_string(static_cast<long long>(n));
      }
      char buf[32];
      std::snprintf(buf, sizeof(buf), "%g", n);
      return std::string(buf);
    }
    case ValueKind::Error:
      return display_name(v.as_error());
    default:
      return v.debug_to_string();
  }
}

// Renders a golden `{kind, value}` cell record (or a bare JSON scalar)
// into the same canonical string `render_value` produces.
std::string render_golden_cell(const JsonValue& cell) {
  if (cell.is_number()) {
    const double n = cell.as_number();
    if (n == static_cast<double>(static_cast<long long>(n))) {
      return std::to_string(static_cast<long long>(n));
    }
    char buf[32];
    std::snprintf(buf, sizeof(buf), "%g", n);
    return std::string(buf);
  }
  if (cell.is_bool()) {
    return cell.as_bool() ? "TRUE" : "FALSE";
  }
  if (cell.is_string()) {
    return cell.as_string();
  }
  if (cell.is_null()) {
    return "";
  }
  if (cell.is_object()) {
    const JsonValue* kind = cell.find("kind");
    const JsonValue* value = cell.find("value");
    if (kind != nullptr && kind->is_string()) {
      if (kind->as_string() == "blank") {
        return "";
      }
      if (kind->as_string() == "error") {
        const JsonValue* code = cell.find("code");
        return code != nullptr && code->is_string() ? code->as_string() : std::string("#ERR");
      }
    }
    if (value != nullptr) {
      return render_golden_cell(*value);
    }
  }
  return "";
}

// ---------------------------------------------------------------------------
// Print pagination comparison
// ---------------------------------------------------------------------------

// Renders a 0-based column index into Excel column letters (0 -> "A").
std::string col_letters(std::uint32_t col) {
  std::string out;
  std::uint32_t n = col + 1U;
  while (n > 0U) {
    const std::uint32_t rem = (n - 1U) % 26U;
    out.insert(out.begin(), static_cast<char>('A' + rem));
    n = (n - 1U) / 26U;
  }
  return out;
}

// Formats one `CellRange` into an A1 range string ("A1:H80"). A
// degenerate single-cell range collapses to a bare "A1".
std::string format_range(const print::CellRange& r) {
  std::string out = col_letters(r.first_col) + std::to_string(r.first_row + 1U);
  if (r.first_row != r.last_row || r.first_col != r.last_col) {
    out += ":";
    out += col_letters(r.last_col) + std::to_string(r.last_row + 1U);
  }
  return out;
}

// Formats a resolved print area (one or more rectangles) into a
// comma-separated A1 string, matching how `expect.print.print_area` is
// authored.
std::string format_print_area(const std::vector<print::CellRange>& ranges) {
  std::string out;
  for (const print::CellRange& r : ranges) {
    if (!out.empty()) {
      out += ",";
    }
    out += format_range(r);
  }
  return out;
}

// Reads a golden integer array (`expect.print.h_breaks` etc.) into a
// vector of 0-based indices.
std::vector<std::uint32_t> golden_index_array(const JsonValue* arr) {
  std::vector<std::uint32_t> out;
  if (arr == nullptr || !arr->is_array()) {
    return out;
  }
  for (const JsonValue& v : arr->as_array()) {
    if (v.is_number()) {
      out.push_back(static_cast<std::uint32_t>(v.as_number()));
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Fixture
// ---------------------------------------------------------------------------

class WorkbookOracleTest : public ::testing::TestWithParam<WorkbookOracleCase> {};

TEST_P(WorkbookOracleTest, Matches) {
  const WorkbookOracleCase& param = GetParam();
  if (param.case_id == "<load-error>") {
    const JsonValue* detail = param.spec.find("error");
    FAIL() << "failed to load " << param.source_file << ": "
           << (detail != nullptr && detail->is_string() ? detail->as_string() : std::string("unknown"));
    return;
  }

  // Divergence-skipped cases land here with a non-empty reason; the
  // golden carries a `"skipped"` field in place of `"expect"`. Mirror
  // the formula track's pattern (tests/oracle/oracle_test.cpp): surface
  // as gtest-skipped so the pass-rate math still reflects them and the
  // intentional gap is visible.
  if (!param.skipped_reason.empty()) {
    GTEST_SKIP() << "divergence.yaml skip-oracle: " << param.skipped_reason;
    return;
  }

  const bool has_pivot = param.spec.find("pivot") != nullptr;
  const bool has_print = param.spec.find("print") != nullptr;

  // A case with neither feature block is nothing this verifier can pin.
  if (!has_pivot && !has_print) {
    GTEST_SKIP() << "case carries no pivot/print block to verify";
    return;
  }

  // --- print path ----------------------------------------------------------
  // Rebuild the workbook from the declarative spec, paginate the named
  // sheet, and diff the `PaginationResult` against `expect.print`.
  if (has_print) {
    auto built_or = build_print_from_spec(param.spec);
    ASSERT_TRUE(static_cast<bool>(built_or)) << "build_print_from_spec failed: " << built_or.error().message;
    const BuiltPrint& built = built_or.value();

    auto pag_or = print::paginate(*built.workbook, built.sheet_index);
    ASSERT_TRUE(static_cast<bool>(pag_or)) << "print::paginate failed: " << pag_or.error().message;
    const print::PaginationResult& pag = pag_or.value();

    const JsonValue* print_expect = param.expect.find("print");
    ASSERT_NE(print_expect, nullptr) << "golden 'expect' has no 'print' block";

    // print_area: exact A1-string equality.
    if (const JsonValue* area_v = print_expect->find("print_area"); area_v != nullptr && area_v->is_string()) {
      EXPECT_EQ(format_print_area(pag.print_area), area_v->as_string()) << "resolved print area mismatch";
    }

    // h_breaks / v_breaks: 1-bit parity with Excel's pagination is
    // best-effort -- font-metric rounding shifts a break by at most one
    // track -- so each expected break position is allowed to differ by
    // +/-1 from the engine's. When the break *counts* match the page
    // count is compared exactly; when they differ (a break landed on a
    // neighbouring track and merged / split a page) the page count is
    // allowed to differ by +/-1.
    const std::vector<std::uint32_t> exp_h = golden_index_array(print_expect->find("h_breaks"));
    const std::vector<std::uint32_t> exp_v = golden_index_array(print_expect->find("v_breaks"));

    auto breaks_within_tolerance = [](const std::vector<std::uint32_t>& got,
                                      const std::vector<std::uint32_t>& want) -> bool {
      if (got.size() != want.size()) {
        return false;
      }
      for (std::size_t i = 0; i < got.size(); ++i) {
        const std::int64_t delta = static_cast<std::int64_t>(got[i]) - static_cast<std::int64_t>(want[i]);
        if (delta < -1 || delta > 1) {
          return false;
        }
      }
      return true;
    };

    auto format_vec = [](const std::vector<std::uint32_t>& vec) -> std::string {
      std::string out = "[";
      for (std::size_t i = 0; i < vec.size(); ++i) {
        if (i > 0) {
          out += ",";
        }
        out += std::to_string(vec[i]);
      }
      out += "]";
      return out;
    };
    const bool h_ok = breaks_within_tolerance(pag.h_breaks, exp_h);
    const bool v_ok = breaks_within_tolerance(pag.v_breaks, exp_v);
    EXPECT_TRUE(h_ok) << "h_breaks differ beyond the +/-1 pagination tolerance: got=" << format_vec(pag.h_breaks)
                      << " want=" << format_vec(exp_h);
    EXPECT_TRUE(v_ok) << "v_breaks differ beyond the +/-1 pagination tolerance: got=" << format_vec(pag.v_breaks)
                      << " want=" << format_vec(exp_v);

    if (const JsonValue* pages_v = print_expect->find("pages"); pages_v != nullptr && pages_v->is_number()) {
      const auto expected_pages = static_cast<std::uint32_t>(pages_v->as_number());
      const bool counts_match = pag.h_breaks.size() == exp_h.size() && pag.v_breaks.size() == exp_v.size();
      if (counts_match) {
        EXPECT_EQ(pag.page_count, expected_pages) << "page count mismatch (break counts agreed)";
      } else {
        const std::int64_t delta =
            static_cast<std::int64_t>(pag.page_count) - static_cast<std::int64_t>(expected_pages);
        EXPECT_GE(delta, -1) << "page count below the +/-1 tolerance";
        EXPECT_LE(delta, 1) << "page count above the +/-1 tolerance";
      }
    }

    if (!has_pivot) {
      return;
    }
  }

  // --- pivot path ----------------------------------------------------------
  // Rebuild the pivot from the declarative spec, evaluate + layout it, and
  // diff the rendered grid (anchor-relative) against `expect.pivot.grid`.
  auto built_or = build_pivot_from_spec(param.spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << "build_pivot_from_spec failed: " << built_or.error().message;
  BuiltPivot built = std::move(built_or.value());

  auto result_or = pivot::evaluate(built.table, built.cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << "pivot::evaluate failed: " << result_or.error().message;

  const pivot::PivotLayoutOptions layout_options = eval::pivot_layout_options_for(built.workbook->excel_profile());
  auto cells_or = pivot::layout(built.table, result_or.value(), layout_options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << "pivot::layout failed: " << cells_or.error().message;
  const pivot::PivotCells& cells = cells_or.value();

  // Flatten the projected cells into an anchor-relative {(r,c) -> value}
  // map. PivotCells coordinates are absolute sheet coords; subtract the
  // anchor to match the golden grid's relative addressing.
  const std::uint32_t anchor_row = built.table.anchor_row();
  const std::uint32_t anchor_col = built.table.anchor_col();
  std::map<std::pair<std::uint32_t, std::uint32_t>, std::string> actual;
  for (const pivot::PivotCell& cell : cells.cells) {
    if (cell.row < anchor_row || cell.col < anchor_col) {
      continue;
    }
    const auto [it, inserted] =
        actual.emplace(std::make_pair(cell.row - anchor_row, cell.col - anchor_col), render_value(cell.value));
    ASSERT_TRUE(inserted) << "duplicate rendered pivot cell at (" << it->first.first << ',' << it->first.second << ')';
  }

  // Extract the golden grid.
  const JsonValue* pivot_expect = param.expect.find("pivot");
  ASSERT_NE(pivot_expect, nullptr) << "golden 'expect' has no 'pivot' block";
  const JsonValue* grid_v = pivot_expect->find("grid");
  ASSERT_NE(grid_v, nullptr) << "golden 'expect.pivot' has no 'grid'";
  ASSERT_TRUE(grid_v->is_array()) << "golden 'expect.pivot.grid' is not an array";

  std::map<std::pair<std::uint32_t, std::uint32_t>, std::string> expected;
  for (const JsonValue& entry : grid_v->as_array()) {
    ASSERT_TRUE(entry.is_object()) << "grid entry is not an object";
    const JsonValue* r_v = entry.find("r");
    const JsonValue* c_v = entry.find("c");
    const JsonValue* val_v = entry.find("value");
    ASSERT_TRUE(r_v != nullptr && r_v->is_number()) << "grid entry missing numeric 'r'";
    ASSERT_TRUE(c_v != nullptr && c_v->is_number()) << "grid entry missing numeric 'c'";
    ASSERT_NE(val_v, nullptr) << "grid entry missing 'value'";

    const auto r = static_cast<std::uint32_t>(r_v->as_number());
    const auto c = static_cast<std::uint32_t>(c_v->as_number());
    const auto [it, inserted] = expected.emplace(std::make_pair(r, c), render_golden_cell(*val_v));
    ASSERT_TRUE(inserted) << "duplicate golden pivot cell at (" << it->first.first << ',' << it->first.second << ')';
  }
  EXPECT_EQ(actual, expected) << "rendered pivot grid differs from the golden grid";

  // A formula probe is evaluated only after the pivot has been attached to
  // the workbook. This is the native counterpart of the Windows driver's
  // post-build GETPIVOTDATA readback and is what pins page/data-axis routing
  // once an external Excel golden is available.
  const JsonValue* pivot_spec = param.spec.find("pivot");
  const JsonValue* probes_v =
      pivot_spec != nullptr && pivot_spec->is_object() ? pivot_spec->find("formula_probes") : nullptr;
  if (probes_v != nullptr) {
    auto probe_results_or = evaluate_pivot_formula_probes(&built, param.spec);
    ASSERT_TRUE(static_cast<bool>(probe_results_or))
        << "evaluate_pivot_formula_probes failed: " << probe_results_or.error().message;
    const JsonValue* expected_probes = param.expect.find("formula_probes");
    ASSERT_NE(expected_probes, nullptr) << "golden 'expect' has no 'formula_probes' block";
    ASSERT_TRUE(expected_probes->is_array()) << "golden 'expect.formula_probes' is not an array";
    ASSERT_EQ(probe_results_or.value().size(), expected_probes->as_array().size())
        << "formula probe result count differs from the golden";
    for (std::size_t i = 0; i < probe_results_or.value().size(); ++i) {
      const FormulaProbeResult& got = probe_results_or.value()[i];
      const JsonValue& want = expected_probes->as_array()[i];
      ASSERT_TRUE(want.is_object()) << "formula probe golden entry is not an object";
      const JsonValue* id_v = want.find("id");
      const JsonValue* result_v = want.find("result");
      ASSERT_NE(id_v, nullptr) << "formula probe golden entry missing 'id'";
      ASSERT_NE(result_v, nullptr) << "formula probe golden entry missing 'result'";
      ASSERT_TRUE(id_v->is_string()) << "formula probe golden id is not a string";
      EXPECT_EQ(got.id, id_v->as_string()) << "formula probe id differs from the golden";
      EXPECT_EQ(render_value(got.value), render_golden_cell(*result_v))
          << "formula probe result differs for " << got.id;
    }
  }
}

// Human-readable gtest parameter names so failures show up as
// `WorkbookOracleTest.Matches/<suite>_<case_id>` instead of a numeric index.
std::string PrintParamName(const ::testing::TestParamInfo<WorkbookOracleCase>& info) {
  std::string name = info.param.suite + "_" + info.param.case_id;
  if (!info.param.variant.empty()) {
    name += "__" + info.param.variant;
  }
  // gtest requires [A-Za-z0-9_]; fold everything else to '_'.
  for (char& c : name) {
    if ((c < 'A' || c > 'Z') && (c < 'a' || c > 'z') && (c < '0' || c > '9') && c != '_') {
      c = '_';
    }
  }
  return name;
}

INSTANTIATE_TEST_SUITE_P(WorkbookOracle, WorkbookOracleTest, ::testing::ValuesIn(workbook_oracle_cases()),
                         PrintParamName);

// With no goldens generated yet the parameter vector is empty and the
// instantiation above expands to nothing. Allow the uninstantiated state
// explicitly so an empty golden_wb tree builds and runs cleanly.
GTEST_ALLOW_UNINSTANTIATED_PARAMETERIZED_TEST(WorkbookOracleTest);

}  // namespace
}  // namespace oracle
}  // namespace tests
}  // namespace formulon
