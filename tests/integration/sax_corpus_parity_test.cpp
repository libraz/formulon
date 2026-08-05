//
// SAX vs DOM parity gate over the 100-book OOXML corpus.
//
// The default OOXML reader path (`io::read_sheet_data`) is a pugixml
// DOM walker. The SAX path (`io::read_sheet_data_sax`) streams cells
// directly off the raw XML. Both paths must produce identical
// Workbook output for every sheet in the corpus.
//
// The corpus generator is shared with `ooxml_corpus_100_test.cpp`
// (see that file for axis decoding); the duplicated helpers are
// minimised to the construction primitives this gate needs to call.
//
// Pipeline per book:
//   1. build wb_a in memory.
//   2. write_ooxml -> bytes_a.
//   3. read_ooxml(bytes_a) using the production reader (which selects
//      DOM or SAX per sheet by size). For all books in this corpus
//      every sheet is well under `kSaxThresholdBytes`, so the
//      production reader picks the DOM path. We capture this read as
//      `wb_dom`.
//   4. Re-walk each `xl/worksheets/sheetN.xml` part directly via the
//      `read_via_sax` adapter so every sheet routes through the SAX
//      path regardless of its size. Capture the result as `wb_sax`.
//   5. Recalc both workbooks.
//   6. Assert the two workbooks are cell-by-cell equal (formula text
//      and cached value).
//
// Tagged with the `SLOW` label because 100 books x two reader passes
// is in the multi-second range under sanitizer builds.

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <deque>
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
#include "io/sheet_reader.h"
#include "io/sst_reader.h"
#include "io/tables_reader.h"
#include "io/workbook_kind.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// ---------------------------------------------------------------------------
// Corpus axes (pruned subset of axes from ooxml_corpus_100_test.cpp).
// We avoid the structured-reference formula axis because the writer
// emits structured refs that the SAX reader treats identically to the
// DOM reader (the SAX path only differs in how XML is parsed; formula
// text passes through verbatim). Sheet shapes stay below
// `kSaxThresholdBytes` (we measure parity on the per-cell decoding
// pipeline, not on the threshold dispatch — the dispatch is exercised
// by the manual SAX-only adapter `read_via_sax`).
// ---------------------------------------------------------------------------

constexpr std::array<std::uint32_t, 4> kSheetCounts = {1U, 2U, 3U, 5U};

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

enum class ValueMix : std::uint8_t {
  kNumbersOnly = 0,
  kNumbersAndText = 1,
  kNumbersTextErrors = 2,
  kNumbersTextErrorsBlanks = 3,
};
constexpr std::array<ValueMix, 4> kValueMixes = {ValueMix::kNumbersOnly, ValueMix::kNumbersAndText,
                                                 ValueMix::kNumbersTextErrors, ValueMix::kNumbersTextErrorsBlanks};

enum class FormulaMode : std::uint8_t {
  kNone = 0,
  kScalarArithmetic = 1,
  kFunctionCalls = 2,
};
constexpr std::array<FormulaMode, 3> kFormulaModes = {FormulaMode::kNone, FormulaMode::kScalarArithmetic,
                                                      FormulaMode::kFunctionCalls};

enum class Decorations : std::uint8_t {
  kNone = 0,
  kDefinedNames = 1,
};
constexpr std::array<Decorations, 2> kDecorations = {Decorations::kNone, Decorations::kDefinedNames};

struct AxisValues {
  std::uint32_t sheet_count;
  CellShape shape;
  ValueMix value_mix;
  FormulaMode formula_mode;
  Decorations decorations;
};

AxisValues axis_values_for(std::uint32_t book_id) {
  // Stride 7 vs product 4*4*4*3*2 = 384; gcd(7, 384) = 1 -> permutation.
  constexpr std::uint32_t kAxisProduct = 4U * 4U * 4U * 3U * 2U;
  constexpr std::uint32_t kStride = 7U;
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

io::ByteSpan span_of(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

std::string describe(std::uint32_t book_id, const AxisValues& a) {
  std::ostringstream os;
  os << "book=" << book_id << " sheets=" << a.sheet_count << " shape=" << a.shape.rows << "x" << a.shape.cols
     << " mix=" << static_cast<int>(a.value_mix) << " formula=" << static_cast<int>(a.formula_mode)
     << " deco=" << static_cast<int>(a.decorations);
  return os.str();
}

// ---------------------------------------------------------------------------
// Workbook construction (mirrors the primary corpus' helpers).
// ---------------------------------------------------------------------------

std::vector<std::string>& text_pool() {
  static std::vector<std::string> pool;
  pool.reserve(1U << 14U);
  return pool;
}

Value pooled_text(const std::string& s) {
  text_pool().push_back(s);
  return Value::text(text_pool().back());
}

Value literal_for(ValueMix mix, std::uint32_t row, std::uint32_t col) {
  const std::uint32_t variants = mix == ValueMix::kNumbersOnly         ? 1U
                                 : mix == ValueMix::kNumbersAndText    ? 2U
                                 : mix == ValueMix::kNumbersTextErrors ? 3U
                                                                       : 4U;
  const std::uint32_t bucket = (row * 7U + col * 3U + 11U) % variants;
  switch (bucket) {
    case 0:
      return Value::number(static_cast<double>(row) * 100.0 + static_cast<double>(col) + 0.5);
    case 1:
      // Use plain ASCII to keep entity coverage on the SAX path
      // exercised separately by the SaxXmlReader unit suite.
      return pooled_text(std::string("t_") + std::to_string(row) + "_" + std::to_string(col));
    case 2:
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

std::string formula_for(FormulaMode mode, std::uint32_t row, std::uint32_t col, std::uint32_t total_rows) {
  if (mode == FormulaMode::kNone) {
    return {};
  }
  if (row == 0U || total_rows < 2U) {
    return {};
  }
  if (((row * 13U + col * 5U) % 4U) != 0U) {
    return {};
  }
  const std::string above = col_letters(col) + std::to_string(row);
  switch (mode) {
    case FormulaMode::kNone:
      return {};
    case FormulaMode::kScalarArithmetic:
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
  }
  return {};
}

Expected<Workbook, Error> build_workbook(std::uint32_t book_id, const AxisValues& a) {
  Workbook wb = Workbook::create_empty();

  for (std::uint32_t s = 0; s < a.sheet_count; ++s) {
    wb.add_sheet(std::string("S") + std::to_string(s + 1U));
  }
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
  for (std::uint32_t s = 1; s < a.sheet_count; ++s) {
    RETURN_IF_ERROR(wb.set_cell_value(s, 0U, 0U, Value::number(static_cast<double>(s) + 0.25)));
  }
  if (a.decorations == Decorations::kDefinedNames) {
    std::vector<io::DefinedName> names;
    io::DefinedName n;
    n.name = std::string("Range_") + std::to_string(book_id);
    n.formula = std::string("S1!$A$1:$") + col_letters(a.shape.cols - 1U) + "$1";
    names.push_back(std::move(n));
    wb.set_defined_names(std::move(names));
  }
  return wb;
}

// ---------------------------------------------------------------------------
// Bypass adapter: drive every sheet through the SAX path regardless of
// size. This is the test-only dual of the production threshold
// dispatch and lets us run the gate without artificially inflating
// each book past 256 KiB.
// ---------------------------------------------------------------------------

/// Re-reads each `xl/worksheets/sheetN.xml` from `bytes` via the SAX
/// reader and overlays the cells onto a freshly constructed Workbook.
/// Defined names / tables / passthrough are not exercised here — those
/// flow through the production reader path unchanged. The caller
/// compares the resulting cells to a DOM-read workbook.
///
/// The shared-string resolution pass is part of the adapter, not an
/// extra: `read_sheet_data_sax` only queues `(row, col, sst_index)` in
/// the context and leaves a `Text("")` placeholder in the cell, exactly
/// as the DOM reader does. Skipping the pass would compare resolved text
/// on one side against placeholders on the other and call the reader
/// paths divergent when they are not.
Expected<Workbook, Error> read_via_sax(io::ByteSpan bytes, const Workbook& template_wb) {
  io::ZipReader zip;
  RETURN_IF_ERROR(zip.open(bytes));

  Workbook wb = Workbook::create_empty();
  for (std::size_t i = 0; i < template_wb.sheet_count(); ++i) {
    wb.add_sheet(std::string(template_wb.sheet(i).name()));
  }

  // The workbook itself owns the text-storage deque now, so the SAX
  // reader appends directly into `wb.mutable_text_storage()`. Cell
  // `Value::text` views remain valid for the workbook's lifetime.
  std::deque<std::string>& text_storage = wb.mutable_text_storage();

  // The corpus writer emits string cells through the shared-string
  // table, so the adapter has to load it before the sheets that index
  // into it. Absent part means the book had no string cells.
  io::SharedStringTable sst;
  if (zip.has_entry("xl/sharedStrings.xml")) {
    auto sst_bytes_or = zip.read_entry("xl/sharedStrings.xml");
    if (!sst_bytes_or) {
      return sst_bytes_or.error();
    }
    auto sst_or = io::read_shared_strings(std::move(sst_bytes_or.value()), text_storage);
    if (!sst_or) {
      return sst_or.error();
    }
    sst = std::move(sst_or.value());
  }

  // Walk the same per-sheet part paths the DOM reader resolves. We
  // enumerate the archive directly because we want the raw bytes
  // regardless of relationship structure; the assumption is that
  // Excel-emitted books name sheets canonically as
  // `xl/worksheets/sheet<i>.xml`. The corpus generator uses exactly
  // this convention, so the adapter does not need full rels parsing.
  for (std::size_t i = 0; i < template_wb.sheet_count(); ++i) {
    std::string path = std::string("xl/worksheets/sheet") + std::to_string(i + 1) + ".xml";
    if (!zip.has_entry(path)) {
      continue;  // tolerated; e.g. binary-only sheet variants we do not care about
    }
    auto sheet_bytes_or = zip.read_entry(path);
    if (!sheet_bytes_or) {
      return sheet_bytes_or.error();
    }
    const std::vector<std::uint8_t>& sheet_bytes = sheet_bytes_or.value();
    io::ByteSpan span{sheet_bytes.data(), sheet_bytes.size()};
    io::SheetReadContext ctx;
    auto rs = io::read_sheet_data_sax(span, i, wb, ctx, text_storage);
    if (!rs) {
      return rs.error();
    }
    for (const auto& [row, col, sst_index] : ctx.pending_sst_cells) {
      if (sst_index >= sst.entries.size()) {
        return make_error(FormulonErrorCode::kIoSheetCorrupt, "shared-string index out of range",
                          "context=sax_corpus_parity sheet_index=" + std::to_string(i));
      }
      wb.sheet(i).set_cell_cached_value_borrowed(row, col, Value::text(sst.entries[sst_index]));
      if (sst_index < sst.phonetic_for_entries.size() && !sst.phonetic_for_entries[sst_index].empty()) {
        wb.sheet(i).set_cell_phonetic(row, col, sst.phonetic_for_entries[sst_index]);
      }
    }
  }
  // Text storage now travels with the workbook itself, so no static
  // anchor is needed: the `Value::text` views remain valid for as
  // long as the returned `Workbook` is alive.
  return wb;
}

// ---------------------------------------------------------------------------
// Comparators (identical contract to ooxml_corpus_100_test.cpp).
// ---------------------------------------------------------------------------

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
  if (a.sheet_count() != b.sheet_count()) {
    return ::testing::AssertionFailure() << "sheet_count differs: " << a.sheet_count() << " vs " << b.sheet_count();
  }
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

// ---------------------------------------------------------------------------
// Parameterised driver.
// ---------------------------------------------------------------------------

constexpr std::uint32_t kCorpusSize = 100U;

class SaxCorpusParity : public ::testing::TestWithParam<std::uint32_t> {};

TEST_P(SaxCorpusParity, DomAndSaxAgree) {
  const std::uint32_t book_id = GetParam();
  const AxisValues axes = axis_values_for(book_id);
  const std::string ctx = describe(book_id, axes);
  SCOPED_TRACE(ctx);

  auto wb_a_or = build_workbook(book_id, axes);
  ASSERT_TRUE(static_cast<bool>(wb_a_or)) << "build failed: " << wb_a_or.error().message << " [" << ctx << "]";
  Workbook& wb_a = wb_a_or.value();

  auto bytes_a_or = wb_a.save();
  ASSERT_TRUE(static_cast<bool>(bytes_a_or)) << "save failed: " << bytes_a_or.error().message << " [" << ctx << "]";

  // DOM read (production reader; sheets are below the SAX threshold).
  auto read_dom_or = io::read_ooxml(span_of(bytes_a_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_dom_or))
      << "read_dom failed: " << read_dom_or.error().message << " [" << ctx << "]";
  Workbook& wb_dom = read_dom_or.value().workbook;
  auto stats_dom_or = wb_dom.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats_dom_or));

  // SAX read (forced).
  auto read_sax_or = read_via_sax(span_of(bytes_a_or.value()), wb_dom);
  ASSERT_TRUE(static_cast<bool>(read_sax_or))
      << "read_sax failed: " << read_sax_or.error().message << " [" << ctx << "]";
  Workbook& wb_sax = read_sax_or.value();
  auto stats_sax_or = wb_sax.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats_sax_or));

  EXPECT_TRUE(cells_match(wb_dom, wb_sax)) << ctx;
}

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

INSTANTIATE_TEST_SUITE_P(Corpus, SaxCorpusParity, ::testing::Range<std::uint32_t>(0U, kCorpusSize),
                         BookIdNameFormatter());

}  // namespace
}  // namespace formulon
