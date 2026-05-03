// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// OOXML round-trip parity over a 48-book programmatically generated
// corpus that focuses on the three subsystems the original 100-book
// corpus deliberately excluded: styles (`xl/styles.xml` + per-cell
// `xf_index`), conditional formats (sheet-level `<conditionalFormatting>`
// blocks), and pivots (`xl/pivotCache/*` plus per-sheet pivot tables).
//
// Sibling to `ooxml_corpus_100_test.cpp`. The 100-book file pins the
// general round-trip surface (cells, formulas, defined names, tables,
// passthrough parts); this file extends the same two-cycle pipeline to
// the styles / CF / pivot surfaces that landed after the 100-book corpus
// was first authored.
//
// Two-cycle pipeline (matches the 100-book file):
//
//   1. construct in-memory (Workbook public API only - no miniz);
//   2. write_ooxml -> bytes_a;
//   3. read_ooxml(bytes_a) -> wb_b;
//   4. wb_b.recalc();
//   5. write_ooxml(wb_b) -> bytes_b;
//   6. read_ooxml(bytes_b) -> wb_c;
//   7. wb_c.recalc();
//   8. assert invariants between wb_b (post-recalc) and wb_c
//      (post-recalc): sheet count + names, per-sheet cell counts and
//      `(row, col, kind, value)` tuples, defined-name list, table
//      list, passthrough-parts list, workbook kind, plus the new
//      surfaces: styles table shape + populated-record equality,
//      conditional-format blocks (sqref + rule list shape), pivot
//      caches and per-sheet pivot tables (count + identity + data
//      fields).
//
// Feature matrix (4 axes):
//
//   * styles    (3): kNone / kBasic / kMixed
//   * cf        (4): kNone / kCellIs / kContainsText / kColorScale
//   * pivot     (2): kNone / kSimple
//   * sheets    (2): 1 / 2
//
// Cartesian product = 3 * 4 * 2 * 2 = 48. The full product is enumerated
// directly (each `BookId in [0, 48)` decodes to one unique combination),
// so the corpus has perfect per-axis coverage with no projection /
// stride. The cell shape is fixed at 5x5: this corpus stresses the new
// auxiliary subsystems, not cell density (the 100-book corpus already
// covers shapes from 1x1 to 100x100).
//
// Same `SLOW` ctest label as `ooxml_corpus_100_test.cpp`: 48 two-cycle
// pipelines through the styles / CF / pivot writer paths is noticeably
// heavier than the default fast suite budget. It runs under
// `ctest -L SLOW` or the full suite, not the default fast pass.

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_types.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "io/ooxml_reader.h"
#include "io/passthrough_part.h"
#include "io/styles_reader.h"
#include "io/tables_reader.h"
#include "io/workbook_kind.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// ---------------------------------------------------------------------------
// Feature matrix.
// ---------------------------------------------------------------------------

/// Styles axis.
enum class StylesAxis : std::uint8_t {
  kNone = 0,   // no styles set; workbook carries the default table only.
  kBasic = 1,  // one custom font (red) on cell A1 via xf=1.
  kMixed = 2,  // three custom CellXfs (varying font colour + fill) on a few cells.
};
constexpr std::array<StylesAxis, 3> kStylesValues = {StylesAxis::kNone, StylesAxis::kBasic, StylesAxis::kMixed};

/// Conditional-format axis.
enum class CfAxis : std::uint8_t {
  kNone = 0,
  kCellIs = 1,        // single `cellIs` GreaterThan 0 rule.
  kContainsText = 2,  // single `containsText` rule with literal "x".
  kColorScale = 3,    // single 2-stop colorScale (Min..Max).
};
constexpr std::array<CfAxis, 4> kCfValues = {CfAxis::kNone, CfAxis::kCellIs, CfAxis::kContainsText,
                                             CfAxis::kColorScale};

/// Pivot axis.
enum class PivotAxis : std::uint8_t {
  kNone = 0,
  kSimple = 1,  // one PivotCache (Region/Amount, three records) + one Sum-of-Amount table.
};
constexpr std::array<PivotAxis, 2> kPivotValues = {PivotAxis::kNone, PivotAxis::kSimple};

/// Sheet-count axis.
constexpr std::array<std::uint32_t, 2> kSheetCounts = {1U, 2U};

/// Cell-shape: fixed at 5x5 across the corpus. The 100-book file stresses
/// shape-density variation; this file stresses the auxiliary subsystems,
/// so the cell payload is small and uniform.
constexpr std::uint32_t kRows = 5U;
constexpr std::uint32_t kCols = 5U;

/// Decoded axis tuple for a corpus book.
struct AxisValues {
  StylesAxis styles;
  CfAxis cf;
  PivotAxis pivot;
  std::uint32_t sheet_count;
};

/// Decodes the axis tuple for book `i`.
///
/// The Cartesian product (3 * 4 * 2 * 2 = 48) is enumerated directly: the
/// caller iterates `i in [0, 48)` and the decoder splits `i` into one
/// position on each axis. With no stride / projection, every axis value
/// appears exactly `48 / |axis|` times, so per-axis coverage is uniform.
AxisValues axis_values_for(std::uint32_t book_id) {
  std::uint32_t v = book_id;
  AxisValues axes{};
  axes.sheet_count = kSheetCounts[v % kSheetCounts.size()];
  v /= static_cast<std::uint32_t>(kSheetCounts.size());
  axes.pivot = kPivotValues[v % kPivotValues.size()];
  v /= static_cast<std::uint32_t>(kPivotValues.size());
  axes.cf = kCfValues[v % kCfValues.size()];
  v /= static_cast<std::uint32_t>(kCfValues.size());
  axes.styles = kStylesValues[v % kStylesValues.size()];
  return axes;
}

/// Human-readable description of an axis tuple. Format:
/// `book=07 styles=basic cf=cellIs pivot=simple sheets=2`.
std::string describe(std::uint32_t book_id, const AxisValues& a) {
  constexpr std::array<const char*, 3> kStylesLabels = {"none", "basic", "mixed"};
  constexpr std::array<const char*, 4> kCfLabels = {"none", "cellIs", "containsText", "colorScale"};
  constexpr std::array<const char*, 2> kPivotLabels = {"none", "simple"};
  std::ostringstream os;
  os << "book=" << book_id << " styles=" << kStylesLabels[static_cast<std::size_t>(a.styles)]
     << " cf=" << kCfLabels[static_cast<std::size_t>(a.cf)]
     << " pivot=" << kPivotLabels[static_cast<std::size_t>(a.pivot)] << " sheets=" << a.sheet_count;
  return os.str();
}

io::ByteSpan span_of(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

// ---------------------------------------------------------------------------
// Workbook construction.
// ---------------------------------------------------------------------------

/// Per-test-process pool for owning text-value strings. Mirrors the
/// 100-book corpus: stable element addresses across the test program's
/// lifetime so `Value::text` (a non-owning view) stays valid through the
/// round-trip pipeline.
std::vector<std::string>& text_pool() {
  static std::vector<std::string> pool;
  pool.reserve(1U << 12U);
  return pool;
}

Value pooled_text(const std::string& s) {
  text_pool().push_back(s);
  return Value::text(text_pool().back());
}

/// Builds a styles table for the `kBasic` / `kMixed` axes.
io::StylesTable build_styles_table(StylesAxis axis) {
  io::StylesTable styles;
  // Indices 0 are the defaults; the corpus pushes one or three custom
  // records. The reader-fallback contract (a missing section yields a
  // single default) means we always populate slot 0 explicitly so the
  // round-trip table layout matches what the writer emits.
  styles.fonts.emplace_back();     // default
  styles.fills.emplace_back();     // default
  styles.borders.emplace_back();   // default
  styles.cell_xfs.emplace_back();  // default

  if (axis == StylesAxis::kNone) {
    return styles;
  }

  // Custom font 1: bold red Calibri.
  io::FontRecord red;
  red.name = "Calibri";
  red.size = 11.0;
  red.bold = true;
  red.color_argb = 0xFFFF0000U;  // ARGB: opaque red.
  styles.fonts.push_back(red);

  // CellXf 1 references font 1.
  io::CellXf xf1;
  xf1.font_index = 1U;
  styles.cell_xfs.push_back(xf1);

  if (axis == StylesAxis::kBasic) {
    return styles;
  }

  // kMixed: one more font, one fill, two more CellXfs.
  io::FontRecord blue;
  blue.name = "Calibri";
  blue.italic = true;
  blue.color_argb = 0xFF0000FFU;
  styles.fonts.push_back(blue);

  io::FillRecord solid_yellow;
  solid_yellow.pattern = 1U;           // solid
  solid_yellow.fg_argb = 0xFFFFFF00U;  // ARGB: opaque yellow.
  styles.fills.push_back(solid_yellow);

  io::CellXf xf2;
  xf2.font_index = 2U;
  xf2.fill_index = 1U;
  styles.cell_xfs.push_back(xf2);

  io::CellXf xf3;
  xf3.font_index = 1U;
  xf3.fill_index = 1U;
  xf3.horizontal_align = 2U;  // center
  styles.cell_xfs.push_back(xf3);

  return styles;
}

/// Builds the conditional-format block list for the `cf` axis.
std::vector<cf::ConditionalFormat> build_conditional_formats(CfAxis axis) {
  std::vector<cf::ConditionalFormat> blocks;
  if (axis == CfAxis::kNone) {
    return blocks;
  }

  cf::ConditionalFormat block{};
  // Three-cell sqref A1:A3 across all the rule kinds the corpus emits.
  block.sqref.push_back({{0U, 0U}, {2U, 0U}});

  cf::CFRule r;
  r.priority = 1;
  switch (axis) {
    case CfAxis::kNone:
      break;
    case CfAxis::kCellIs:
      r.type = cf::RuleType::CellIs;
      r.op = cf::CellIsOperator::GreaterThan;
      r.formula1 = "0";
      break;
    case CfAxis::kContainsText:
      r.type = cf::RuleType::ContainsText;
      r.text = "x";
      break;
    case CfAxis::kColorScale: {
      r.type = cf::RuleType::ColorScale;
      cf::ColorScaleSpec spec;
      spec.thresholds.push_back({cf::CfvoType::Min, "", true});
      spec.thresholds.push_back({cf::CfvoType::Max, "", true});
      spec.colors.push_back({255, 0, 0, 255});  // red
      spec.colors.push_back({0, 255, 0, 255});  // green
      r.color_scale = std::move(spec);
      break;
    }
  }
  block.rules.push_back(std::move(r));
  blocks.push_back(std::move(block));
  return blocks;
}

/// Convenience: appends `s` to `cache.mutable_text_storage()` and returns
/// a `Value::text` aliasing that storage. Mirrors `MakeText` in
/// `ooxml_writer_pivot_test.cpp`; the cache owns the bytes for the
/// lifetime of the `PivotCache`.
Value cache_text(pivot::PivotCache& cache, std::string s) {
  cache.mutable_text_storage().push_back(std::move(s));
  return Value::text(cache.mutable_text_storage().back());
}

/// Builds a minimal Region/Amount pivot cache (cache_id=0, three records:
/// North/100, South/200, North/300). Identical shape to the canonical
/// cache used in `ooxml_writer_pivot_test.cpp`, intentionally so the
/// writer / reader contract is exercised against the same payload.
pivot::PivotCache build_pivot_cache() {
  pivot::PivotCache cache;
  cache.set_cache_id(0U);

  pivot::PivotCacheField region;
  region.name = "Region";
  region.shared_items.push_back(cache_text(cache, "North"));
  region.shared_items.push_back(cache_text(cache, "South"));
  cache.mutable_fields().push_back(std::move(region));

  pivot::PivotCacheField amount;
  amount.name = "Amount";
  cache.mutable_fields().push_back(std::move(amount));

  pivot::PivotCacheRecord r0;
  r0.cells.push_back(Value::number(0.0));  // -> "North"
  r0.cells.push_back(Value::number(100.0));
  cache.mutable_records().push_back(std::move(r0));
  pivot::PivotCacheRecord r1;
  r1.cells.push_back(Value::number(1.0));  // -> "South"
  r1.cells.push_back(Value::number(200.0));
  cache.mutable_records().push_back(std::move(r1));
  pivot::PivotCacheRecord r2;
  r2.cells.push_back(Value::number(0.0));  // -> "North"
  r2.cells.push_back(Value::number(300.0));
  cache.mutable_records().push_back(std::move(r2));

  return cache;
}

/// Builds a Sum-of-Amount-by-Region pivot table anchored at
/// `(anchor_row, anchor_col)`, bound to cache id 0.
std::unique_ptr<pivot::PivotTable> build_pivot_table(std::uint32_t anchor_row, std::uint32_t anchor_col) {
  auto table = std::make_unique<pivot::PivotTable>();
  table->set_name("PivotTable1");
  table->set_pivot_cache_id(0U);
  table->set_anchor(anchor_row, anchor_col, 5U, 2U);

  pivot::PivotField region;
  region.axis = pivot::PivotAxis::Row;
  region.custom_name = "Region";
  region.items.push_back(pivot::PivotItem{"", true});
  region.items.push_back(pivot::PivotItem{"", true});
  table->mutable_fields().push_back(std::move(region));

  pivot::PivotField amount;
  amount.axis = pivot::PivotAxis::Value;
  amount.custom_name = "Amount";
  table->mutable_fields().push_back(std::move(amount));

  table->mutable_row_field_order().push_back(0U);

  pivot::PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 1U;
  sum_amount.aggregation = pivot::Aggregation::Sum;
  table->mutable_data_fields().push_back(std::move(sum_amount));
  return table;
}

/// Builds the corpus workbook for `book_id` according to its axis tuple.
Expected<Workbook, Error> build_workbook(std::uint32_t book_id, const AxisValues& a) {
  static_cast<void>(book_id);

  Workbook wb = Workbook::create_empty();

  // Sheets. The pivot table is anchored on the last sheet when present,
  // mirroring how Excel typically lays out a pivot on its own worksheet.
  for (std::uint32_t s = 0; s < a.sheet_count; ++s) {
    wb.add_sheet(std::string("S") + std::to_string(s + 1U));
  }

  // Populate the first sheet with a small 5x5 numeric grid and one
  // string cell so the `containsText` CF rule (and the writer's text
  // path) has something to match.
  for (std::uint32_t r = 0; r < kRows; ++r) {
    for (std::uint32_t c = 0; c < kCols; ++c) {
      // Distinct numeric pattern so equality is meaningful.
      RETURN_IF_ERROR(wb.set_cell_value(0U, r, c, Value::number(static_cast<double>(r * kCols + c) + 0.25)));
    }
  }
  // One literal text cell at A4 (row=3, col=0) carrying "x_marker"; the
  // ContainsText axis searches for "x" so this cell will match.
  RETURN_IF_ERROR(wb.set_cell_value(0U, 3U, 0U, pooled_text("x_marker")));

  // Secondary sheets get one literal each so multi-sheet round-trip
  // still drags every sheet through the writer / reader path.
  for (std::uint32_t s = 1; s < a.sheet_count; ++s) {
    RETURN_IF_ERROR(wb.set_cell_value(s, 0U, 0U, Value::number(static_cast<double>(s) + 0.5)));
  }

  // Styles. `set_styles` replaces the table wholesale; `set_cell_xf_index`
  // applies the index to a cell that already exists.
  if (a.styles != StylesAxis::kNone) {
    wb.set_styles(build_styles_table(a.styles));
    // Always style A1 with xf=1.
    RETURN_IF_ERROR(wb.set_cell_xf_index(0U, 0U, 0U, 1U));
    if (a.styles == StylesAxis::kMixed) {
      // A2 with xf=2 (italic blue + yellow fill), B1 with xf=3 (bold red
      // + yellow fill, centre-aligned). Both cells were populated above
      // so the index applies cleanly.
      RETURN_IF_ERROR(wb.set_cell_xf_index(0U, 1U, 0U, 2U));
      RETURN_IF_ERROR(wb.set_cell_xf_index(0U, 0U, 1U, 3U));
    }
  }

  // Conditional formats. Stored on the first sheet.
  if (a.cf != CfAxis::kNone) {
    wb.sheet(0).mutable_conditional_formats() = build_conditional_formats(a.cf);
  }

  // Pivot cache + table. Anchored on the last sheet (sheet index
  // `sheet_count - 1`) at D1; a 1-sheet workbook puts both the source
  // grid and the pivot anchor on the same sheet, which the reader/writer
  // accept (the pivot reads from the cache, not the live cells).
  if (a.pivot == PivotAxis::kSimple) {
    wb.add_pivot_cache(std::make_unique<pivot::PivotCache>(build_pivot_cache()));
    const std::uint32_t anchor_sheet = a.sheet_count - 1U;
    wb.sheet(anchor_sheet).add_pivot_table(build_pivot_table(/*anchor_row=*/0U, /*anchor_col=*/3U));
  }

  return wb;
}

// ---------------------------------------------------------------------------
// Cross-cycle invariant comparators. The contract mirrors the 100-book
// corpus: `wb_b` (post-recalc-1) and `wb_c` (post-recalc-2) must match on
// every on-disk-observable surface the reader populates.
// ---------------------------------------------------------------------------

::testing::AssertionResult sheet_shapes_match(const Workbook& a, const Workbook& b) {
  if (a.sheet_count() != b.sheet_count()) {
    return ::testing::AssertionFailure() << "sheet_count differs: " << a.sheet_count() << " vs " << b.sheet_count();
  }
  for (std::size_t i = 0; i < a.sheet_count(); ++i) {
    if (a.sheet(i).name() != b.sheet(i).name()) {
      return ::testing::AssertionFailure()
             << "sheet[" << i << "] name differs: '" << a.sheet(i).name() << "' vs '" << b.sheet(i).name() << "'";
    }
    if (a.sheet(i).cell_count() != b.sheet(i).cell_count()) {
      return ::testing::AssertionFailure() << "sheet[" << i << "] cell_count differs: " << a.sheet(i).cell_count()
                                           << " vs " << b.sheet(i).cell_count();
    }
  }
  return ::testing::AssertionSuccess();
}

bool values_equal(const Value& x, const Value& y) {
  if (x.kind() != y.kind()) {
    return false;
  }
  switch (x.kind()) {
    case ValueKind::Blank:
      return true;
    case ValueKind::Number:
      return x.as_number() == y.as_number();
    case ValueKind::Bool:
      return x.as_boolean() == y.as_boolean();
    case ValueKind::Text:
      return x.as_text() == y.as_text();
    case ValueKind::Error:
      return x.as_error() == y.as_error();
    default:
      return false;
  }
}

::testing::AssertionResult cells_match(const Workbook& a, const Workbook& b) {
  for (std::size_t s = 0; s < a.sheet_count(); ++s) {
    const Sheet& sa = a.sheet(s);
    const Sheet& sb = b.sheet(s);
    for (const auto& [row, cells] : sa.rows()) {
      for (std::uint32_t col = 0; col < cells.size(); ++col) {
        const Cell& ca = cells[col];
        const bool a_has_data = !ca.formula_text.empty() || !ca.cached_value.is_blank();
        if (!a_has_data) {
          continue;
        }
        const Cell* cb = sb.cell_at(row, col);
        if (cb == nullptr) {
          return ::testing::AssertionFailure() << "sheet[" << s << "] (" << row << "," << col << ") missing on side B";
        }
        if (ca.formula_text != cb->formula_text) {
          return ::testing::AssertionFailure() << "sheet[" << s << "] (" << row << "," << col << ") formula differs: '"
                                               << ca.formula_text << "' vs '" << cb->formula_text << "'";
        }
        if (!values_equal(ca.cached_value, cb->cached_value)) {
          return ::testing::AssertionFailure() << "sheet[" << s << "] (" << row << "," << col
                                               << ") value differs: kind_a=" << static_cast<int>(ca.cached_value.kind())
                                               << " kind_b=" << static_cast<int>(cb->cached_value.kind());
        }
      }
    }
    for (const auto& [row, cells] : sb.rows()) {
      for (std::uint32_t col = 0; col < cells.size(); ++col) {
        const Cell& cb = cells[col];
        const bool b_has_data = !cb.formula_text.empty() || !cb.cached_value.is_blank();
        if (!b_has_data) {
          continue;
        }
        const Cell* ca = sa.cell_at(row, col);
        if (ca == nullptr) {
          return ::testing::AssertionFailure() << "sheet[" << s << "] (" << row << "," << col << ") missing on side A";
        }
      }
    }
  }
  return ::testing::AssertionSuccess();
}

::testing::AssertionResult defined_names_match(const Workbook& a, const Workbook& b) {
  if (a.defined_names().size() != b.defined_names().size()) {
    return ::testing::AssertionFailure() << "defined_names.size differs: " << a.defined_names().size() << " vs "
                                         << b.defined_names().size();
  }
  for (std::size_t i = 0; i < a.defined_names().size(); ++i) {
    const io::DefinedName& x = a.defined_names()[i];
    const io::DefinedName& y = b.defined_names()[i];
    if (x.name != y.name || x.formula != y.formula || x.local_sheet_id != y.local_sheet_id) {
      return ::testing::AssertionFailure()
             << "defined_names[" << i << "] differs: name='" << x.name << "' vs '" << y.name << "', formula='"
             << x.formula << "' vs '" << y.formula << "', scope=" << x.local_sheet_id << " vs " << y.local_sheet_id;
    }
  }
  return ::testing::AssertionSuccess();
}

::testing::AssertionResult tables_match(const Workbook& a, const Workbook& b) {
  if (a.tables().size() != b.tables().size()) {
    return ::testing::AssertionFailure() << "tables.size differs: " << a.tables().size() << " vs " << b.tables().size();
  }
  for (std::size_t i = 0; i < a.tables().size(); ++i) {
    const io::TableMetadata& x = a.tables()[i];
    const io::TableMetadata& y = b.tables()[i];
    if (x.name != y.name || x.ref != y.ref || x.sheet_index != y.sheet_index || x.columns.size() != y.columns.size()) {
      return ::testing::AssertionFailure() << "tables[" << i << "] structural mismatch";
    }
  }
  return ::testing::AssertionSuccess();
}

::testing::AssertionResult passthrough_match(const Workbook& a, const Workbook& b) {
  if (a.passthrough_parts().size() != b.passthrough_parts().size()) {
    return ::testing::AssertionFailure() << "passthrough_parts.size differs: " << a.passthrough_parts().size() << " vs "
                                         << b.passthrough_parts().size();
  }
  for (const io::PassthroughPart& x : a.passthrough_parts()) {
    auto it = std::find_if(b.passthrough_parts().begin(), b.passthrough_parts().end(),
                           [&x](const io::PassthroughPart& y) { return y.path == x.path; });
    if (it == b.passthrough_parts().end()) {
      return ::testing::AssertionFailure() << "passthrough '" << x.path << "' missing on side B";
    }
    if (x.bytes.size() != it->bytes.size() || !std::equal(x.bytes.begin(), x.bytes.end(), it->bytes.begin())) {
      return ::testing::AssertionFailure() << "passthrough '" << x.path << "' bytes differ";
    }
  }
  return ::testing::AssertionSuccess();
}

// ---------------------------------------------------------------------------
// Styles invariants.
//
// We compare:
//   * top-level section sizes (`fonts`, `fills`, `borders`, `num_fmts`,
//     `cell_xfs`) so any record drift surfaces immediately;
//   * deep equality on the populated slices the corpus actually uses
//     (font name / colour / bold / italic; fill pattern / fg_argb;
//     CellXf font/fill/border/num_fmt/alignment fields).
// ---------------------------------------------------------------------------

::testing::AssertionResult fonts_equal_at(const io::FontRecord& x, const io::FontRecord& y, std::size_t i) {
  if (x.name != y.name || x.size != y.size || x.bold != y.bold || x.italic != y.italic || x.strike != y.strike ||
      x.underline != y.underline || x.color_argb != y.color_argb) {
    return ::testing::AssertionFailure() << "fonts[" << i << "] differs";
  }
  return ::testing::AssertionSuccess();
}

::testing::AssertionResult fills_equal_at(const io::FillRecord& x, const io::FillRecord& y, std::size_t i) {
  if (x.pattern != y.pattern || x.fg_argb != y.fg_argb || x.bg_argb != y.bg_argb) {
    return ::testing::AssertionFailure() << "fills[" << i << "] differs";
  }
  return ::testing::AssertionSuccess();
}

::testing::AssertionResult cell_xfs_equal_at(const io::CellXf& x, const io::CellXf& y, std::size_t i) {
  if (x.font_index != y.font_index || x.fill_index != y.fill_index || x.border_index != y.border_index ||
      x.num_fmt_id != y.num_fmt_id || x.horizontal_align != y.horizontal_align ||
      x.vertical_align != y.vertical_align || x.wrap_text != y.wrap_text) {
    return ::testing::AssertionFailure() << "cell_xfs[" << i << "] differs";
  }
  return ::testing::AssertionSuccess();
}

::testing::AssertionResult styles_match(const Workbook& a, const Workbook& b) {
  const io::StylesTable& sa = a.styles();
  const io::StylesTable& sb = b.styles();
  if (sa.fonts.size() != sb.fonts.size()) {
    return ::testing::AssertionFailure() << "styles.fonts.size differs: " << sa.fonts.size() << " vs "
                                         << sb.fonts.size();
  }
  if (sa.fills.size() != sb.fills.size()) {
    return ::testing::AssertionFailure() << "styles.fills.size differs: " << sa.fills.size() << " vs "
                                         << sb.fills.size();
  }
  if (sa.borders.size() != sb.borders.size()) {
    return ::testing::AssertionFailure() << "styles.borders.size differs: " << sa.borders.size() << " vs "
                                         << sb.borders.size();
  }
  if (sa.num_fmts.size() != sb.num_fmts.size()) {
    return ::testing::AssertionFailure() << "styles.num_fmts.size differs: " << sa.num_fmts.size() << " vs "
                                         << sb.num_fmts.size();
  }
  if (sa.cell_xfs.size() != sb.cell_xfs.size()) {
    return ::testing::AssertionFailure() << "styles.cell_xfs.size differs: " << sa.cell_xfs.size() << " vs "
                                         << sb.cell_xfs.size();
  }
  // Deep equality on every populated slot. The reader's empty-section
  // fallback means even kNone corpora carry one default record per
  // section, which still compares equal under this loop.
  for (std::size_t i = 0; i < sa.fonts.size(); ++i) {
    auto r = fonts_equal_at(sa.fonts[i], sb.fonts[i], i);
    if (!r) {
      return r;
    }
  }
  for (std::size_t i = 0; i < sa.fills.size(); ++i) {
    auto r = fills_equal_at(sa.fills[i], sb.fills[i], i);
    if (!r) {
      return r;
    }
  }
  for (std::size_t i = 0; i < sa.cell_xfs.size(); ++i) {
    auto r = cell_xfs_equal_at(sa.cell_xfs[i], sb.cell_xfs[i], i);
    if (!r) {
      return r;
    }
  }
  return ::testing::AssertionSuccess();
}

// ---------------------------------------------------------------------------
// Conditional-format invariants. We compare:
//   * per-sheet block count;
//   * per-block sqref count;
//   * per-rule type, priority, op, formula1, text — the fields the
//     corpus emits across its four CF axes.
// ---------------------------------------------------------------------------

::testing::AssertionResult cf_match(const Workbook& a, const Workbook& b) {
  for (std::size_t s = 0; s < a.sheet_count(); ++s) {
    const auto& cfs_a = a.sheet(s).conditional_formats();
    const auto& cfs_b = b.sheet(s).conditional_formats();
    if (cfs_a.size() != cfs_b.size()) {
      return ::testing::AssertionFailure()
             << "sheet[" << s << "] cf.size differs: " << cfs_a.size() << " vs " << cfs_b.size();
    }
    for (std::size_t i = 0; i < cfs_a.size(); ++i) {
      const auto& ba = cfs_a[i];
      const auto& bb = cfs_b[i];
      if (ba.sqref.size() != bb.sqref.size()) {
        return ::testing::AssertionFailure()
               << "sheet[" << s << "] cf[" << i << "] sqref.size differs: " << ba.sqref.size() << " vs "
               << bb.sqref.size();
      }
      if (ba.rules.size() != bb.rules.size()) {
        return ::testing::AssertionFailure()
               << "sheet[" << s << "] cf[" << i << "] rules.size differs: " << ba.rules.size() << " vs "
               << bb.rules.size();
      }
      for (std::size_t k = 0; k < ba.rules.size(); ++k) {
        const auto& ra = ba.rules[k];
        const auto& rb = bb.rules[k];
        if (ra.type != rb.type) {
          return ::testing::AssertionFailure()
                 << "sheet[" << s << "] cf[" << i << "] rules[" << k << "].type differs: " << static_cast<int>(ra.type)
                 << " vs " << static_cast<int>(rb.type);
        }
        if (ra.priority != rb.priority) {
          return ::testing::AssertionFailure() << "sheet[" << s << "] cf[" << i << "] rules[" << k
                                               << "].priority differs: " << ra.priority << " vs " << rb.priority;
        }
        if (ra.op.has_value() != rb.op.has_value() || (ra.op.has_value() && ra.op.value() != rb.op.value())) {
          return ::testing::AssertionFailure() << "sheet[" << s << "] cf[" << i << "] rules[" << k << "].op differs";
        }
        if (ra.formula1 != rb.formula1) {
          return ::testing::AssertionFailure()
                 << "sheet[" << s << "] cf[" << i << "] rules[" << k << "].formula1 differs";
        }
        if (ra.text != rb.text) {
          return ::testing::AssertionFailure() << "sheet[" << s << "] cf[" << i << "] rules[" << k << "].text differs";
        }
      }
    }
  }
  return ::testing::AssertionSuccess();
}

// ---------------------------------------------------------------------------
// Pivot invariants. We compare:
//   * workbook-level pivot-cache count + per-cache cache_id and field
//     count;
//   * per-sheet pivot-table count, name, anchor, and data-field count.
// The full record-by-record cache equality is already covered by
// `ooxml_writer_pivot_test.cpp`; this corpus asserts the structural
// shape survives across the two-cycle pipeline.
// ---------------------------------------------------------------------------

::testing::AssertionResult pivot_match(const Workbook& a, const Workbook& b) {
  if (a.pivot_caches().size() != b.pivot_caches().size()) {
    return ::testing::AssertionFailure() << "pivot_caches.size differs: " << a.pivot_caches().size() << " vs "
                                         << b.pivot_caches().size();
  }
  for (std::size_t i = 0; i < a.pivot_caches().size(); ++i) {
    const pivot::PivotCache* ca = a.pivot_caches()[i].get();
    const pivot::PivotCache* cb = b.pivot_caches()[i].get();
    if (ca == nullptr || cb == nullptr) {
      return ::testing::AssertionFailure() << "pivot_caches[" << i << "] null";
    }
    if (ca->cache_id() != cb->cache_id()) {
      return ::testing::AssertionFailure()
             << "pivot_caches[" << i << "] cache_id differs: " << ca->cache_id() << " vs " << cb->cache_id();
    }
    if (ca->fields().size() != cb->fields().size()) {
      return ::testing::AssertionFailure() << "pivot_caches[" << i << "] fields.size differs: " << ca->fields().size()
                                           << " vs " << cb->fields().size();
    }
    if (ca->records().size() != cb->records().size()) {
      return ::testing::AssertionFailure() << "pivot_caches[" << i << "] records.size differs: " << ca->records().size()
                                           << " vs " << cb->records().size();
    }
  }
  for (std::size_t s = 0; s < a.sheet_count(); ++s) {
    const auto& tas = a.sheet(s).pivot_tables();
    const auto& tbs = b.sheet(s).pivot_tables();
    if (tas.size() != tbs.size()) {
      return ::testing::AssertionFailure()
             << "sheet[" << s << "] pivot_tables.size differs: " << tas.size() << " vs " << tbs.size();
    }
    for (std::size_t i = 0; i < tas.size(); ++i) {
      const pivot::PivotTable* x = tas[i].get();
      const pivot::PivotTable* y = tbs[i].get();
      if (x == nullptr || y == nullptr) {
        return ::testing::AssertionFailure() << "sheet[" << s << "] pivot_tables[" << i << "] null";
      }
      if (x->name() != y->name()) {
        return ::testing::AssertionFailure() << "sheet[" << s << "] pivot_tables[" << i << "] name differs: '"
                                             << x->name() << "' vs '" << y->name() << "'";
      }
      if (x->anchor_row() != y->anchor_row() || x->anchor_col() != y->anchor_col()) {
        return ::testing::AssertionFailure() << "sheet[" << s << "] pivot_tables[" << i << "] anchor differs";
      }
      if (x->data_fields().size() != y->data_fields().size()) {
        return ::testing::AssertionFailure()
               << "sheet[" << s << "] pivot_tables[" << i << "] data_fields.size differs: " << x->data_fields().size()
               << " vs " << y->data_fields().size();
      }
    }
  }
  return ::testing::AssertionSuccess();
}

// ---------------------------------------------------------------------------
// Parameterised driver.
// ---------------------------------------------------------------------------

constexpr std::uint32_t kCorpusSize = 48U;

class OoxmlCorpusExtended : public ::testing::TestWithParam<std::uint32_t> {};

TEST_P(OoxmlCorpusExtended, TwoCyclePipeline) {
  const std::uint32_t book_id = GetParam();
  const AxisValues axes = axis_values_for(book_id);
  const std::string ctx = describe(book_id, axes);
  SCOPED_TRACE(ctx);

  // (1) Build wb_a.
  auto wb_a_or = build_workbook(book_id, axes);
  ASSERT_TRUE(static_cast<bool>(wb_a_or)) << "build failed: " << wb_a_or.error().message << " [" << ctx << "]";
  Workbook& wb_a = wb_a_or.value();

  // (2) Write -> bytes_a.
  auto bytes_a_or = wb_a.save();
  ASSERT_TRUE(static_cast<bool>(bytes_a_or)) << "save_a failed: " << bytes_a_or.error().message << " [" << ctx << "]";

  // (3) Read bytes_a -> wb_b.
  auto read_b_or = io::read_ooxml(span_of(bytes_a_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_b_or)) << "read_b failed: " << read_b_or.error().message << " [" << ctx << "]";
  Workbook& wb_b = read_b_or.value().workbook;

  // (4) Recalc wb_b.
  auto stats_b_or = wb_b.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats_b_or)) << "recalc_b failed: " << stats_b_or.error().message << " [" << ctx << "]";

  // (5) Write wb_b -> bytes_b.
  auto bytes_b_or = wb_b.save();
  ASSERT_TRUE(static_cast<bool>(bytes_b_or)) << "save_b failed: " << bytes_b_or.error().message << " [" << ctx << "]";

  // (6) Read bytes_b -> wb_c.
  auto read_c_or = io::read_ooxml(span_of(bytes_b_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_c_or)) << "read_c failed: " << read_c_or.error().message << " [" << ctx << "]";
  Workbook& wb_c = read_c_or.value().workbook;

  // (7) Recalc wb_c.
  auto stats_c_or = wb_c.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats_c_or)) << "recalc_c failed: " << stats_c_or.error().message << " [" << ctx << "]";

  // (8) Cross-cycle invariants.
  EXPECT_TRUE(sheet_shapes_match(wb_b, wb_c)) << ctx;
  EXPECT_TRUE(cells_match(wb_b, wb_c)) << ctx;
  EXPECT_TRUE(defined_names_match(wb_b, wb_c)) << ctx;
  EXPECT_TRUE(tables_match(wb_b, wb_c)) << ctx;
  EXPECT_TRUE(passthrough_match(wb_b, wb_c)) << ctx;
  EXPECT_TRUE(styles_match(wb_b, wb_c)) << ctx;
  EXPECT_TRUE(cf_match(wb_b, wb_c)) << ctx;
  EXPECT_TRUE(pivot_match(wb_b, wb_c)) << ctx;
  EXPECT_EQ(wb_b.kind(), wb_c.kind()) << ctx;
}

/// GTest parameter formatter: each test case is named `Book00` ... `Book47`
/// so failures sort lexicographically and the failing index is immediately
/// legible. Two-digit padding is enough for a 48-book corpus.
struct BookIdNameFormatter {
  std::string operator()(const ::testing::TestParamInfo<std::uint32_t>& info) const {
    std::ostringstream os;
    os << "Book";
    if (info.param < 10U) {
      os << '0';
    }
    os << info.param;
    return os.str();
  }
};

INSTANTIATE_TEST_SUITE_P(Corpus, OoxmlCorpusExtended, ::testing::Range<std::uint32_t>(0U, kCorpusSize),
                         BookIdNameFormatter());

}  // namespace
}  // namespace formulon
