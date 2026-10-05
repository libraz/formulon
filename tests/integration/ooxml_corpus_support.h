#pragma once

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <functional>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon::ooxml_corpus {

inline io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

/// `save()` wrapper that ASSERTs success and returns the bytes.
inline Expected<std::vector<std::uint8_t>, Error> SaveBytes(const Workbook& wb) {
  return wb.save();
}

/// Two workbooks have the same sheet shape iff they have the same
/// number of sheets, the i-th sheet on each has the same display name,
/// and the i-th sheet on each holds the same total `cell_count()`. We
/// intentionally do NOT compare the geometric `(rows x cols)` extents
/// because the row store grows lazily and the on-disk projection
/// depends only on populated cells.
inline bool sheets_have_same_shape(const Workbook& a, const Workbook& b) {
  if (a.sheet_count() != b.sheet_count()) {
    return false;
  }
  for (std::size_t i = 0; i < a.sheet_count(); ++i) {
    if (a.sheet(i).name() != b.sheet(i).name()) {
      return false;
    }
    if (a.sheet(i).cell_count() != b.sheet(i).cell_count()) {
      return false;
    }
  }
  return true;
}

/// Order-preserving by name, plus the four scalar fields. The reader
/// preserves declaration order, and the writer emits in that order, so
/// equality must be index-aligned rather than set-equal.
inline bool defined_names_equal(const Workbook& a, const Workbook& b) {
  if (a.defined_names().size() != b.defined_names().size()) {
    return false;
  }
  for (std::size_t i = 0; i < a.defined_names().size(); ++i) {
    const DefinedName& x = a.defined_names()[i];
    const DefinedName& y = b.defined_names()[i];
    if (x.name != y.name || x.formula != y.formula || x.local_sheet_id != y.local_sheet_id || x.hidden != y.hidden ||
        x.comment != y.comment) {
      return false;
    }
  }
  return true;
}

/// Tables compared by id, ref, name, display_name, header/totals row
/// flags, sheet index, and column-list shape (id + name). Other column
/// fields (totals_label / totals_function) are not asserted here
/// because the corpus does not stress them across all books; the
/// per-book invariants check those where relevant.
inline bool tables_equal(const Workbook& a, const Workbook& b) {
  if (a.tables().size() != b.tables().size()) {
    return false;
  }
  for (std::size_t i = 0; i < a.tables().size(); ++i) {
    const TableMetadata& x = a.tables()[i];
    const TableMetadata& y = b.tables()[i];
    if (x.id != y.id || x.name != y.name || x.display_name != y.display_name || x.ref != y.ref ||
        x.sheet_index != y.sheet_index || x.header_row != y.header_row || x.totals_row != y.totals_row) {
      return false;
    }
    if (x.columns.size() != y.columns.size()) {
      return false;
    }
    for (std::size_t c = 0; c < x.columns.size(); ++c) {
      if (x.columns[c].id != y.columns[c].id || x.columns[c].name != y.columns[c].name) {
        return false;
      }
    }
  }
  return true;
}

/// Passthrough parts are compared by path + raw bytes (content_type
/// is also compared when both sides have one). Order is preserved in
/// the round-trip but the assertion is set-style: we look up by path
/// so re-emission ordering is not load-bearing.
inline bool passthrough_parts_equal(const Workbook& a, const Workbook& b) {
  if (a.passthrough_parts().size() != b.passthrough_parts().size()) {
    return false;
  }
  for (const PassthroughPart& x : a.passthrough_parts()) {
    auto it = std::find_if(b.passthrough_parts().begin(), b.passthrough_parts().end(),
                           [&x](const PassthroughPart& y) { return y.path == x.path; });
    if (it == b.passthrough_parts().end()) {
      return false;
    }
    if (x.bytes.size() != it->bytes.size()) {
      return false;
    }
    if (!std::equal(x.bytes.begin(), x.bytes.end(), it->bytes.begin())) {
      return false;
    }
    if (!x.content_type.empty() && !it->content_type.empty() && x.content_type != it->content_type) {
      return false;
    }
  }
  return true;
}

/// Iterates every cell with a non-empty formula on `a` and asserts the
/// same coordinate on `b` carries the same formula text. Used by the
/// volatile / iterative cases where cached values legitimately drift
/// across recalcs but the formulas themselves must round-trip
/// verbatim.
inline bool formula_texts_equal_for_each_cell(const Workbook& a, const Workbook& b) {
  if (a.sheet_count() != b.sheet_count()) {
    return false;
  }
  for (std::size_t s = 0; s < a.sheet_count(); ++s) {
    const Sheet& sa = a.sheet(s);
    const Sheet& sb = b.sheet(s);
    for (const auto& [row, cells] : sa.rows()) {
      for (std::uint32_t col = 0; col < cells.size(); ++col) {
        const Cell& ca = cells[col];
        if (ca.formula_text.empty()) {
          continue;
        }
        const Cell* cb = sb.cell_at(row, col);
        if (cb == nullptr || cb->formula_text != ca.formula_text) {
          return false;
        }
      }
    }
  }
  return true;
}

struct CorpusBook {
  std::string id;
  std::function<Expected<std::vector<std::uint8_t>, Error>()> build;
  std::function<void(const Workbook&)> assert_invariants;
};

Expected<std::vector<std::uint8_t>, Error> BuildPassthroughThemeArchive();
Expected<std::vector<std::uint8_t>, Error> BuildSstArchive();

Expected<std::vector<std::uint8_t>, Error> BuildEmpty();
Expected<std::vector<std::uint8_t>, Error> BuildSingleLiteral();
Expected<std::vector<std::uint8_t>, Error> BuildLiteralsAllKinds();
Expected<std::vector<std::uint8_t>, Error> BuildArithmetic();
Expected<std::vector<std::uint8_t>, Error> BuildCrossSheet();
Expected<std::vector<std::uint8_t>, Error> BuildMultiSheetIndependent();
Expected<std::vector<std::uint8_t>, Error> BuildUnicodeSheetNames();
Expected<std::vector<std::uint8_t>, Error> BuildUnicodeCellText();
Expected<std::vector<std::uint8_t>, Error> BuildRangeAggregates();
Expected<std::vector<std::uint8_t>, Error> BuildVolatileNowToday();
Expected<std::vector<std::uint8_t>, Error> BuildIterativeCircular();
Expected<std::vector<std::uint8_t>, Error> BuildDefinedNamesWorkbookScope();
Expected<std::vector<std::uint8_t>, Error> BuildDefinedNamesSheetScope();
Expected<std::vector<std::uint8_t>, Error> BuildSingleTable();
Expected<std::vector<std::uint8_t>, Error> BuildMultipleTables();
Expected<std::vector<std::uint8_t>, Error> BuildNestedFunctionCalls();
Expected<std::vector<std::uint8_t>, Error> BuildLargeGrid500Cells();
Expected<std::vector<std::uint8_t>, Error> BuildKitchenSink();

void AssertEmpty(const Workbook& wb);
void AssertSingleLiteral(const Workbook& wb);
void AssertLiteralsAllKinds(const Workbook& wb);
void AssertArithmetic(const Workbook& wb);
void AssertCrossSheet(const Workbook& wb);
void AssertMultiSheetIndependent(const Workbook& wb);
void AssertUnicodeSheetNames(const Workbook& wb);
void AssertUnicodeCellText(const Workbook& wb);
void AssertRangeAggregates(const Workbook& wb);
void AssertVolatileNowToday(const Workbook& wb);
void AssertIterativeCircular(const Workbook& wb);
void AssertDefinedNamesWorkbookScope(const Workbook& wb);
void AssertDefinedNamesSheetScope(const Workbook& wb);
void AssertSingleTable(const Workbook& wb);
void AssertMultipleTables(const Workbook& wb);
void AssertNestedFunctionCalls(const Workbook& wb);
void AssertLargeGrid(const Workbook& wb);
void AssertPassthroughTheme(const Workbook& wb);
void AssertSstCells(const Workbook& wb);
void AssertKitchenSink(const Workbook& wb);

std::vector<CorpusBook> make_corpus();

}  // namespace formulon::ooxml_corpus
