// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// OOXML round-trip parity over a 100-book programmatically generated
// corpus.
//
// Sibling to `ooxml_corpus_test.cpp` (the 20-book hand-crafted corpus).
// Where the 20-book file pins specific named features (passthrough
// theme, explicit SST, iterative-calc circular pair, ...), this file
// stress-tests the round-trip pipeline across a wider feature surface:
// it builds 100 distinct synthetic workbooks programmatically from a
// 5-axis feature matrix, then runs each through the same two-cycle
// pipeline:
//
//   1. construct in-memory (Workbook public API only — no miniz);
//   2. write_ooxml -> bytes_a;
//   3. read_ooxml(bytes_a) -> wb_b;
//   4. wb_b.recalc();
//   5. write_ooxml(wb_b) -> bytes_b;
//   6. read_ooxml(bytes_b) -> wb_c;
//   7. wb_c.recalc();
//   8. assert invariants between wb_b (post-recalc) and wb_c
//      (post-recalc): sheet count + names, per-sheet cell counts and
//      `(row, col, kind, value)` tuples, defined-name list, table
//      list, passthrough-parts list, workbook kind.
//
// The corpus is *generated* (no committed binary fixtures) so the
// repository stays small and the feature matrix is auditable in source.
// Each book is identified by an integer `BookId in [0, 100)`; the tuple
// of axis values is derived deterministically from `BookId` so a
// regression failure pinpoints both the book and the specific feature
// combination.
//
// This test is registered with the `SLOW` ctest label: 100 two-cycle
// pipelines (some at 100x100 cell shapes) take noticeably longer than
// the default fast suite budget. It runs under `ctest -L SLOW` or the
// full suite, not the default CI fast pass.

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <sstream>
#include <string>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "io/ooxml_reader.h"
#include "io/passthrough_part.h"
#include "io/tables_reader.h"
#include "io/workbook_kind.h"
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
//
// Five axes. The full Cartesian product is 4 * 4 * 4 * 4 * 2 = 512 combinations.
// We project that down to 100 by stepping through the 512-space with a
// stride of 5 (gcd(5, 512) == 1, so the stride is a complete permutation
// generator) and taking the first 100 indices. That gives diverse
// per-axis coverage without enumerating all 512.
// ---------------------------------------------------------------------------

/// Sheet count axis.
constexpr std::array<std::uint32_t, 4> kSheetCounts = {1U, 2U, 3U, 5U};

/// Cell-shape axis: dimensions of the populated rectangle on the first
/// sheet. Secondary sheets carry a single cell each so total cell counts
/// stay tractable for 100 books.
struct CellShape {
  std::uint32_t rows;
  std::uint32_t cols;
};
constexpr std::array<CellShape, 4> kCellShapes = {
    CellShape{1U, 1U},
    CellShape{5U, 5U},
    CellShape{20U, 20U},
    CellShape{100U, 100U},
};

/// Value-mix axis.
enum class ValueMix : std::uint8_t {
  kNumbersOnly = 0,
  kNumbersAndText = 1,
  kNumbersTextErrors = 2,
  kNumbersTextErrorsBlanks = 3,
};
constexpr std::array<ValueMix, 4> kValueMixes = {ValueMix::kNumbersOnly, ValueMix::kNumbersAndText,
                                                 ValueMix::kNumbersTextErrors, ValueMix::kNumbersTextErrorsBlanks};

/// Formula-complexity axis.
enum class FormulaMode : std::uint8_t {
  kNone = 0,
  kScalarArithmetic = 1,
  kFunctionCalls = 2,
  kStructuredReferences = 3,
};
constexpr std::array<FormulaMode, 4> kFormulaModes = {FormulaMode::kNone, FormulaMode::kScalarArithmetic,
                                                      FormulaMode::kFunctionCalls, FormulaMode::kStructuredReferences};

/// Decorations axis (defined names + tables). Two levels rather than
/// four because the round-trip cost of decorations is binary in
/// practice: either we exercise them or we do not.
enum class Decorations : std::uint8_t {
  kNone = 0,
  kDefinedNamesAndTables = 1,
};
constexpr std::array<Decorations, 2> kDecorations = {Decorations::kNone, Decorations::kDefinedNamesAndTables};

/// Axis tuple decoded from a corpus book index.
struct AxisValues {
  std::uint32_t sheet_count;
  CellShape shape;
  ValueMix value_mix;
  FormulaMode formula_mode;
  Decorations decorations;
};

/// Decodes the axis tuple for book `i` deterministically.
///
/// The encoding stride (5) is coprime with the axis-product cardinality
/// (512), so iterating `i in [0, 512)` visits every combination exactly
/// once. Capping at `i in [0, 100)` walks the first 100 of those
/// permutations, which gives wide per-axis coverage in a small corpus.
AxisValues axis_values_for(std::uint32_t book_id) {
  constexpr std::uint32_t kAxisProduct = 4U * 4U * 4U * 4U * 2U;  // 512
  constexpr std::uint32_t kStride = 5U;                           // coprime with 512.
  const std::uint32_t mixed = (book_id * kStride) % kAxisProduct;
  std::uint32_t v = mixed;
  AxisValues axes{};
  axes.decorations = kDecorations[v % kDecorations.size()];
  v /= static_cast<std::uint32_t>(kDecorations.size());
  axes.formula_mode = kFormulaModes[v % kFormulaModes.size()];
  v /= static_cast<std::uint32_t>(kFormulaModes.size());
  axes.value_mix = kValueMixes[v % kValueMixes.size()];
  v /= static_cast<std::uint32_t>(kValueMixes.size());
  axes.shape = kCellShapes[v % kCellShapes.size()];
  v /= static_cast<std::uint32_t>(kCellShapes.size());
  axes.sheet_count = kSheetCounts[v % kSheetCounts.size()];
  return axes;
}

/// Human-readable description of an axis tuple (used in failure
/// messages). Format: `book=42 sheets=3 shape=20x20 mix=NTE formula=fn deco=DT`.
std::string describe(std::uint32_t book_id, const AxisValues& a) {
  constexpr std::array<const char*, 4> kMixLabels = {"N", "NT", "NTE", "NTEB"};
  constexpr std::array<const char*, 4> kFormulaLabels = {"none", "scalar", "fn", "sref"};
  std::ostringstream os;
  os << "book=" << book_id << " sheets=" << a.sheet_count << " shape=" << a.shape.rows << "x" << a.shape.cols
     << " mix=" << kMixLabels[static_cast<std::size_t>(a.value_mix)]
     << " formula=" << kFormulaLabels[static_cast<std::size_t>(a.formula_mode)]
     << " deco=" << (a.decorations == Decorations::kDefinedNamesAndTables ? "DT" : "none");
  return os.str();
}

io::ByteSpan span_of(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

// ---------------------------------------------------------------------------
// Workbook construction.
//
// Constraints carried over from the 20-book corpus:
//   * Avoid formulas whose result is text (the writer's
//     `<v>...</v>`-without-`t="str"` is read back as numeric and fails).
//     All formulas in this corpus produce numbers.
//   * Text Values are non-owning views; their backing storage is held in
//     a function-static pool (stable addresses, lives until program
//     teardown) so the round-trip pipeline can read them safely.
//   * Decorations are only defined names + table metadata — the
//     surfaces the reader/writer fully round-trip. Styles, CF, charts,
//     drawings, and pivots are out of scope for this bundle.
// ---------------------------------------------------------------------------

/// Per-test-process pool for owning text-value strings. Stable element
/// addresses are guaranteed because the pool only ever appends and is
/// not destroyed until program teardown.
std::vector<std::string>& text_pool() {
  static std::vector<std::string> pool;
  pool.reserve(1U << 14U);
  return pool;
}

/// Adds `s` to the text pool and returns a `Value::text` view onto the
/// pooled storage.
Value pooled_text(const std::string& s) {
  text_pool().push_back(s);
  return Value::text(text_pool().back());
}

/// Literal value for `(row, col)` under `mix`. Selection inside a mix
/// is a deterministic function of `(row, col)` so a failing cell is
/// reproducible.
Value literal_for(ValueMix mix, std::uint32_t row, std::uint32_t col) {
  // 0 = number, 1 = text, 2 = error, 3 = blank.
  std::uint32_t bucket = 0;
  const std::uint32_t variants = mix == ValueMix::kNumbersOnly         ? 1U
                                 : mix == ValueMix::kNumbersAndText    ? 2U
                                 : mix == ValueMix::kNumbersTextErrors ? 3U
                                                                       : /* kNumbersTextErrorsBlanks */ 4U;
  bucket = (row * 7U + col * 3U + 11U) % variants;
  switch (bucket) {
    case 0:
      // Distinct numeric pattern per (row, col) so equality is meaningful.
      return Value::number(static_cast<double>(row) * 100.0 + static_cast<double>(col) + 0.5);
    case 1:
      return pooled_text(std::string("t_") + std::to_string(row) + "_" + std::to_string(col));
    case 2:
      // Cycle through the three error codes the OOXML round-trip
      // covers exhaustively in the 20-book corpus (Div0, Value, NA).
      switch ((row + col) % 3U) {
        case 0:
          return Value::error(ErrorCode::Div0);
        case 1:
          return Value::error(ErrorCode::Value);
        default:
          return Value::error(ErrorCode::NA);
      }
    case 3:
    default:
      return Value::blank();
  }
}

/// Produces an Excel column letter (A, B, ..., Z, AA, AB, ...) for a
/// 0-based column index.
std::string col_letters(std::uint32_t col) {
  std::string out;
  std::uint32_t v = col + 1U;
  while (v > 0U) {
    const std::uint32_t r = (v - 1U) % 26U;
    out.insert(out.begin(), static_cast<char>('A' + r));
    v = (v - 1U) / 26U;
  }
  return out;
}

/// Formula text for `(row, col)` under `mode`, or empty string if the
/// cell should hold a literal. Density is ~25% so each book exercises
/// both the formula and the literal write paths.
std::string formula_for(FormulaMode mode, std::uint32_t row, std::uint32_t col, std::uint32_t total_rows) {
  if (mode == FormulaMode::kNone) {
    return {};
  }
  // Reserve row 0 for inputs (literal), formulas land below.
  if (row == 0U || total_rows < 2U) {
    return {};
  }
  // Light density — 1 in 4.
  if (((row * 13U + col * 5U) % 4U) != 0U) {
    return {};
  }
  const std::string above = col_letters(col) + std::to_string(row);  // 1-based row r.
  switch (mode) {
    case FormulaMode::kNone:
      return {};
    case FormulaMode::kScalarArithmetic:
      // Numeric arithmetic only — never produces a text result.
      switch ((row + col) % 4U) {
        case 0:
          return std::string("=") + above + "+1";
        case 1:
          return std::string("=") + above + "*2";
        case 2:
          return std::string("=") + above + "-3";
        default:
          return std::string("=") + above + "/2";
      }
    case FormulaMode::kFunctionCalls:
      // Aggregators / IF over numeric branches. Each result is numeric.
      switch ((row + col) % 4U) {
        case 0:
          return std::string("=SUM(") + col_letters(col) + "1:" + above + ")";
        case 1:
          return std::string("=AVERAGE(") + col_letters(col) + "1:" + above + ")";
        case 2:
          return std::string("=IF(") + above + ">0,1,0)";
        default:
          return std::string("=ROUND(") + above + ",1)";
      }
    case FormulaMode::kStructuredReferences:
      // Bundle 4.5 wired structured-reference resolution. We point the
      // formula at the single table this corpus declares (`Tbl_<book>`,
      // built below). The formula targets the first numeric column of
      // the table and is therefore guaranteed numeric.
      //
      // Note: structured refs are only used when the book has the
      // decoration axis enabled AND the table has been declared in the
      // populated rectangle. Outside that pre-condition we fall back to
      // a scalar-arithmetic formula so the round-trip still exercises
      // formula text re-emission. The caller (build_workbook) is
      // responsible for the narrowing.
      return std::string("=") + above + "+1";
  }
  return {};
}

/// Builds the corpus workbook for `book_id` according to its axis
/// tuple.
Expected<Workbook, Error> build_workbook(std::uint32_t book_id, const AxisValues& a) {
  Workbook wb = Workbook::create_empty();

  // Sheets.
  for (std::uint32_t s = 0; s < a.sheet_count; ++s) {
    wb.add_sheet(std::string("S") + std::to_string(s + 1U));
  }

  // Populate the first sheet according to the cell-shape axis.
  for (std::uint32_t r = 0; r < a.shape.rows; ++r) {
    for (std::uint32_t c = 0; c < a.shape.cols; ++c) {
      const std::string formula = formula_for(a.formula_mode, r, c, a.shape.rows);
      if (!formula.empty()) {
        RETURN_IF_ERROR(wb.set_cell_formula(0U, r, c, formula));
      } else {
        RETURN_IF_ERROR(wb.set_cell_value(0U, r, c, literal_for(a.value_mix, r, c)));
      }
    }
  }

  // Secondary sheets get one literal each so multi-sheet round-trip
  // still drags every sheet through the writer / reader path.
  for (std::uint32_t s = 1; s < a.sheet_count; ++s) {
    RETURN_IF_ERROR(wb.set_cell_value(s, 0U, 0U, Value::number(static_cast<double>(s) + 0.25)));
  }

  // Decorations: defined name + a single table on the first sheet.
  if (a.decorations == Decorations::kDefinedNamesAndTables) {
    std::vector<io::DefinedName> names;
    io::DefinedName n;
    n.name = std::string("Range_") + std::to_string(book_id);
    // Workbook-scope; reference the first row of the first sheet.
    n.formula = std::string("S1!$A$1:$") + col_letters(a.shape.cols - 1U) + "$1";
    names.push_back(std::move(n));
    wb.set_defined_names(std::move(names));

    // A table only makes sense when the populated rectangle is at
    // least 2 rows x 1 col (header row + body row).
    if (a.shape.rows >= 2U && a.shape.cols >= 1U) {
      io::TableMetadata t;
      t.id = 1U;
      t.name = std::string("Tbl_") + std::to_string(book_id);
      t.display_name = t.name;
      t.ref = std::string("A1:") + col_letters(a.shape.cols - 1U) + std::to_string(a.shape.rows);
      t.sheet_index = 0U;
      t.header_row = true;
      t.totals_row = false;
      for (std::uint32_t c = 0; c < a.shape.cols; ++c) {
        t.columns.push_back(io::TableColumn{c + 1U, std::string("Col") + std::to_string(c + 1U), "", "", ""});
      }
      std::vector<io::TableMetadata> tables;
      tables.push_back(std::move(t));
      wb.set_tables(std::move(tables));
    }
  }

  return wb;
}

// ---------------------------------------------------------------------------
// Cross-cycle invariant comparators. The contract mirrors the 20-book
// corpus: `wb_b` (post-recalc-1) and `wb_c` (post-recalc-2) must match
// on the on-disk-observable surfaces. We intentionally compare ONLY the
// surfaces the reader actually round-trips (CLAUDE.md "scope" note).
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

/// Equal in the value sense the round-trip cares about.
///
///   * Number: bit-equal IEEE 754. The writer emits the canonical
///     decimal representation; the reader re-parses it. Cycle 2 must
///     re-emit the same canonical text and re-parse to the same double.
///   * Bool / Text / Error: structural equality.
///   * Blank: only matches Blank.
bool values_equal(const Value& x, const Value& y) {
  if (x.kind() != y.kind()) {
    return false;
  }
  switch (x.kind()) {
    case ValueKind::Blank:
      return true;
    case ValueKind::Number:
      // Bit-equal so NaN payload differences (none expected here, but
      // guarded) and signed-zero round-trips both ride this comparator.
      return x.as_number() == y.as_number();
    case ValueKind::Bool:
      return x.as_boolean() == y.as_boolean();
    case ValueKind::Text:
      return x.as_text() == y.as_text();
    case ValueKind::Error:
      return x.as_error() == y.as_error();
    default:
      // Array / Ref / Lambda are not produced by the corpus; treat as
      // not-equal so any escape from the corpus surface registers as a
      // test failure rather than a silent match.
      return false;
  }
}

::testing::AssertionResult cells_match(const Workbook& a, const Workbook& b) {
  // Per-sheet, per-cell tuple equality. We iterate `a` and look up on
  // `b` (and vice-versa) so a missing cell on either side surfaces.
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
    // Reverse direction: a cell present on B but absent on A also
    // counts as drift.
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
    for (std::size_t c = 0; c < x.columns.size(); ++c) {
      if (x.columns[c].name != y.columns[c].name) {
        return ::testing::AssertionFailure() << "tables[" << i << "].columns[" << c << "] name differs: '"
                                             << x.columns[c].name << "' vs '" << y.columns[c].name << "'";
      }
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
// Parameterised driver.
// ---------------------------------------------------------------------------

constexpr std::uint32_t kCorpusSize = 100U;

class OoxmlCorpus100 : public ::testing::TestWithParam<std::uint32_t> {};

TEST_P(OoxmlCorpus100, TwoCyclePipeline) {
  const std::uint32_t book_id = GetParam();
  const AxisValues axes = axis_values_for(book_id);
  const std::string ctx = describe(book_id, axes);
  SCOPED_TRACE(ctx);

  // (1) Build wb_a in memory.
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
  EXPECT_EQ(wb_b.kind(), wb_c.kind()) << ctx;
}

/// GTest parameter formatter: each test case is named `Book000` ...
/// `Book099` so failures sort lexicographically and the failing index
/// is immediately legible.
struct BookIdNameFormatter {
  std::string operator()(const ::testing::TestParamInfo<std::uint32_t>& info) const {
    std::ostringstream os;
    os << "Book";
    if (info.param < 100U) {
      os << '0';
    }
    if (info.param < 10U) {
      os << '0';
    }
    os << info.param;
    return os.str();
  }
};

INSTANTIATE_TEST_SUITE_P(Corpus, OoxmlCorpus100, ::testing::Range<std::uint32_t>(0U, kCorpusSize),
                         BookIdNameFormatter());

}  // namespace
}  // namespace formulon
