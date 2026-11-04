//
// `Workbook::project_pivot` against the report cells Excel saved.
//
// Every pivot of the Excel-saved fixtures under `pivot_{layout,options,
// aggregation,show_as}/` is loaded, its range is overwritten, and the
// projection must write back what Excel wrote: the value and its type, and the
// cell format's `pivotButton`, horizontal alignment and indent, and applied
// number format. Vertical alignment is not compared: Excel adds
// `vertical="center"` to every pivot cell format of a session, which is not a
// property of the pivot.
//
// The cases below the fixture sweep cover what a single load cannot: the
// merges `mergeItem` adds, a projection after rows were inserted above the
// report, pending projections, and a refused overlap.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <filesystem>
#include <map>
#include <memory>
#include <optional>
#include <set>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "merge_range.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "styles.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

/// `relative` under the Excel fixture directory, e.g. "pivot_layout/x.xlsx".
std::unique_ptr<io::OoxmlReadResult> ReadFixture(const std::string& relative) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + relative;
  const std::vector<std::uint8_t> bytes = ReadFileBytes(path);
  EXPECT_FALSE(bytes.empty()) << "could not read " << path;
  auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << relative << ": " << (read_or ? "" : read_or.error().message);
  if (!read_or) {
    return nullptr;
  }
  return std::make_unique<io::OoxmlReadResult>(std::move(read_or.value()));
}

/// The UI locale Excel laid a fixture out in.
ExcelProfile FixtureProfile(const std::string& relative) {
  ExcelProfile profile;
  const bool english = relative.rfind("pivot_layout/en_", 0) == 0;
  EXPECT_TRUE(parse_excel_profile_id(english ? "mac-365-en_US" : "mac-365-ja_JP", &profile));
  return profile;
}

std::string ValueText(const Value& v) {
  if (v.is_blank()) {
    return "";
  }
  if (v.is_text()) {
    return "s:" + std::string(v.as_text());
  }
  if (v.is_number()) {
    char buf[40];
    std::snprintf(buf, sizeof(buf), "n:%.17g", v.as_number());
    return buf;
  }
  if (v.is_boolean()) {
    return v.as_boolean() ? "b:1" : "b:0";
  }
  if (v.is_error()) {
    return "e:" + std::to_string(static_cast<int>(v.as_error()));
  }
  return "?";
}

/// The format facts Excel gives a pivot cell, from its `cellXfs` record.
std::string FormatText(const StylesTable& styles, std::uint32_t xf_index) {
  if (xf_index >= styles.cell_xfs.size()) {
    return "xf?";
  }
  const CellXf& xf = styles.cell_xfs[xf_index];
  std::string out;
  if (xf.extra_attrs.find("pivotButton=\"1\"") != std::string::npos) {
    out += " button";
  }
  if (HasHorizontalAlign(xf) && xf.horizontal_align != 0U) {
    out += " h" + std::to_string(xf.horizontal_align);
  }
  if (xf.has_indent && xf.indent != 0U) {
    out += " indent" + std::to_string(xf.indent);
  }
  if (xf.apply_number_format) {
    out += " nf" + std::to_string(xf.num_fmt_id);
  }
  return out;
}

/// What a sheet coordinate holds, or nothing for no cell.
struct CellFacts {
  bool present = false;
  std::string value;
  std::string format;
};

CellFacts Facts(const Workbook& wb, const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  CellFacts facts;
  const Cell* cell = sheet.cell_at(row, col);
  if (cell == nullptr) {
    return facts;
  }
  facts.value = ValueText(cell->cached_value);
  facts.format = FormatText(wb.styles(), cell->xf_index);
  facts.present = !facts.value.empty() || !facts.format.empty();
  return facts;
}

std::map<std::pair<std::uint32_t, std::uint32_t>, CellFacts> Snapshot(const Workbook& wb, const Sheet& sheet) {
  std::map<std::pair<std::uint32_t, std::uint32_t>, CellFacts> out;
  for (const auto& [row, cells] : sheet.rows()) {
    for (const RowCells::Entry entry : cells) {
      out[{row, entry.col}] = Facts(wb, sheet, row, entry.col);
    }
  }
  return out;
}

/// The range a loaded pivot covers: its location, and the page fields plus
/// their gap row above it.
MergeRange LoadedExtent(const pivot::PivotTable& table) {
  MergeRange r{table.anchor_row(), table.anchor_col(), table.anchor_row() + table.span_rows() - 1U,
               table.anchor_col() + table.span_cols() - 1U};
  const std::uint32_t pages = table.location_row_page_count().value_or(0U);
  if (pages > 0U) {
    r.first_row -= pages + 1U;
    const std::uint32_t width = table.location_col_page_count().value_or(1U) * 3U - 1U;
    r.last_col = std::max(r.last_col, table.anchor_col() + width - 1U);
  }
  return r;
}

MergeRange Union(const MergeRange& a, const MergeRange& b) {
  return MergeRange{std::min(a.first_row, b.first_row), std::min(a.first_col, b.first_col),
                    std::max(a.last_row, b.last_row), std::max(a.last_col, b.last_col)};
}

std::string A1(std::uint32_t row, std::uint32_t col) {
  std::string letters;
  for (std::uint32_t c = col + 1U; c > 0U; c = (c - 1U) / 26U) {
    letters.insert(letters.begin(), static_cast<char>('A' + (c - 1U) % 26U));
  }
  return letters + std::to_string(row + 1U);
}

// ---------------------------------------------------------------------------
// Excel's own report cells
// ---------------------------------------------------------------------------

/// Which cells of a fixture's pivot are left out of the comparison, and why.
struct Exclusions {
  /// Value cells Excel wrote without a format although their siblings carry
  /// the data field's: only their format is skipped.
  bool data_format = false;
  /// Header rows whose cells Excel wrote differently from the sibling
  /// fixtures of the same shape: skipped whole.
  bool header_rows = false;
  /// One cell whose format alone departs from its siblings'.
  std::optional<std::pair<std::uint32_t, std::uint32_t>> format_cell;
};

Exclusions ExclusionsFor(const std::string& relative) {
  Exclusions ex;
  const std::string name = relative.substr(relative.find('/') + 1U);
  // Data cells with no `s` at all, with `numFmtId` 0 either way.
  for (const char* file : {"ja_d13_values_only.xlsx", "show_details_leaf.xlsx", "show_details_twice.xlsx",
                           "show_details_subtotal_and_grand.xlsx", "date_field_ungrouped.xlsx"}) {
    ex.data_format = ex.data_format || name == file;
  }
  // Multi-level column headers whose row-label caption and corner cells are
  // left unwritten, and date reports whose header row sits at the location's
  // first row although `firstHeaderRow` is 1.
  ex.header_rows = name == "K4__average.xlsx" || name == "L22__var.xlsx" || name == "date_field_ungrouped.xlsx" ||
                   name.rfind("date_group_", 0) == 0;
  // A data-field label at its own position written unindented.
  if (name == "ja_d5_values_first_rows.xlsx") {
    ex.format_cell = std::make_pair(7U, 0U);  // A8
  }
  return ex;
}

/// Every fixture file of the swept directories, sorted.
std::vector<std::string> FixtureFiles() {
  std::set<std::string> files;
  for (const char* dir : {"pivot_layout", "pivot_options", "pivot_aggregation", "pivot_show_as"}) {
    std::error_code ec;
    const std::filesystem::path root = std::filesystem::path(FORMULON_FIXTURES_DIR) / "excel" / dir;
    for (std::filesystem::directory_iterator it(root, ec), end; !ec && it != end; it.increment(ec)) {
      if (it->path().extension() == ".xlsx") {
        files.insert(std::string(dir) + "/" + it->path().filename().string());
      }
    }
    EXPECT_FALSE(ec) << dir << ": " << ec.message();
  }
  return std::vector<std::string>(files.begin(), files.end());
}

TEST(PivotProjectionFixture, EveryPivotProjectsTheCellsExcelSaved) {
  std::size_t pivots = 0;
  for (const std::string& relative : FixtureFiles()) {
    // Saved without its records, the cache cannot be evaluated again.
    if (relative == "pivot_options/op_savedata_off.xlsx") {
      continue;
    }
    std::unique_ptr<io::OoxmlReadResult> read = ReadFixture(relative);
    if (read == nullptr) {
      continue;
    }
    Workbook& wb = read->workbook;
    wb.set_excel_profile(FixtureProfile(relative));
    const Exclusions ex = ExclusionsFor(relative);
    for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
      for (std::size_t p = 0; p < wb.sheet(s).pivot_tables().size(); ++p) {
        const pivot::PivotTable& table = *wb.sheet(s).pivot_tables()[p];
        const MergeRange loaded = LoadedExtent(table);
        const std::uint32_t data_row = table.anchor_row() + table.location_first_data_row().value_or(1U);
        const std::uint32_t data_col = table.anchor_col() + table.location_first_data_col().value_or(1U);
        const auto excel = Snapshot(wb, wb.sheet(s));
        // Stale content the projection has to replace.
        for (std::uint32_t r = loaded.first_row; r <= loaded.last_row; ++r) {
          for (std::uint32_t c = loaded.first_col; c <= loaded.last_col; ++c) {
            ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(s, r, c, "stale")));
          }
        }
        auto projected_or = wb.project_pivot(s, p);
        ASSERT_TRUE(static_cast<bool>(projected_or)) << relative << ": " << projected_or.error().message;
        ++pivots;
        ASSERT_TRUE(table.projected_extent().has_value()) << relative;
        const pivot::PivotRect& e = *table.projected_extent();
        const MergeRange region = Union(loaded, MergeRange{e.row0, e.col0, e.row1, e.col1});
        std::ostringstream diff;
        for (std::uint32_t r = region.first_row; r <= region.last_row; ++r) {
          if (ex.header_rows && r >= table.anchor_row() && r < data_row) {
            continue;
          }
          for (std::uint32_t c = region.first_col; c <= region.last_col; ++c) {
            const auto it = excel.find({r, c});
            const CellFacts want = it == excel.end() ? CellFacts{} : it->second;
            CellFacts got = Facts(wb, wb.sheet(s), r, c);
            const bool skip_format = (ex.data_format && r >= data_row && c >= data_col) ||
                                     (ex.format_cell.has_value() && *ex.format_cell == std::make_pair(r, c));
            if (skip_format) {
              got.format = want.format;
              got.present = !got.value.empty() || !got.format.empty() || want.present;
            }
            if (got.present != want.present || got.value != want.value || got.format != want.format) {
              diff << "  " << A1(r, c)
                   << ": excel=" << (want.present ? "[" + want.value + "|" + want.format + "]" : "-")
                   << " engine=" << (got.present ? "[" + got.value + "|" + got.format + "]" : "-") << "\n";
            }
          }
        }
        EXPECT_TRUE(diff.str().empty()) << relative << " sheet " << wb.sheet(s).name() << " pivot " << table.name()
                                        << "\n"
                                        << diff.str();
      }
    }
  }
  RecordProperty("pivots", static_cast<int>(pivots));
  std::printf("[ projection ] %zu pivots compared\n", pivots);
  EXPECT_GT(pivots, 0U);
}

// ---------------------------------------------------------------------------
// Merges, structural edits, pending projections, refusals
// ---------------------------------------------------------------------------

/// The fixture's only pivot sheet.
std::size_t PivotSheet(const Workbook& wb) {
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    if (!wb.sheet(s).pivot_tables().empty()) {
      return s;
    }
  }
  ADD_FAILURE() << "no pivot";
  return 0;
}

std::set<std::string> MergeTexts(const Sheet& sheet) {
  std::set<std::string> out;
  for (const MergeRange& m : sheet.merges()) {
    out.insert(A1(m.first_row, m.first_col) + ":" + A1(m.last_row, m.last_col));
  }
  return out;
}

TEST(PivotProjectionFixture, MergeItemAddsTheMergesExcelSaved) {
  std::unique_ptr<io::OoxmlReadResult> read = ReadFixture("pivot_options/op_merge_tabular.xlsx");
  ASSERT_NE(read, nullptr);
  Workbook& wb = read->workbook;
  wb.set_excel_profile(FixtureProfile("pivot_options/op_merge_tabular.xlsx"));
  const std::size_t s = PivotSheet(wb);
  const std::set<std::string> excel = MergeTexts(wb.sheet(s));
  ASSERT_EQ(excel, (std::set<std::string>{"A5:B5", "A6:A7", "A8:B8", "A9:A10", "A11:B11", "A12:B12"}));
  ASSERT_TRUE(static_cast<bool>(wb.clear_merges(s)));
  ASSERT_TRUE(static_cast<bool>(wb.project_pivot(s, 0)));
  EXPECT_EQ(MergeTexts(wb.sheet(s)), excel);
  EXPECT_EQ(wb.sheet(s).pivot_tables()[0]->projected_merges().size(), 6U);
  // Projecting again replaces the merges rather than stacking them.
  ASSERT_TRUE(static_cast<bool>(wb.project_pivot(s, 0)));
  EXPECT_EQ(wb.sheet(s).merges().size(), 6U);
}

TEST(PivotProjectionFixture, ReprojectionAfterARowInsertClearsTheMovedReport) {
  std::unique_ptr<io::OoxmlReadResult> read = ReadFixture("pivot_layout/ja_c0_rows.xlsx");
  ASSERT_NE(read, nullptr);
  Workbook& wb = read->workbook;
  wb.set_excel_profile(FixtureProfile("pivot_layout/ja_c0_rows.xlsx"));
  const std::size_t s = PivotSheet(wb);
  ASSERT_TRUE(static_cast<bool>(wb.project_pivot(s, 0)));
  pivot::PivotTable& table = *wb.sheet(s).mutable_pivot_tables()[0];
  const pivot::PivotRect before = *table.projected_extent();
  const auto excel = Snapshot(wb, wb.sheet(s));

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(s, 0, 2)));
  ASSERT_TRUE(table.projected_extent().has_value());
  const pivot::PivotRect moved = *table.projected_extent();
  EXPECT_EQ(moved.row0, before.row0 + 2U);
  EXPECT_EQ(moved.row1, before.row1 + 2U);
  EXPECT_EQ(table.anchor_row(), before.row0 + 2U);

  // Without its grand-total row the report is one row shorter; the row it
  // left at the moved position must be cleared, not the one at the old one.
  table.set_grand_totals(table.grand_totals_rows(), false);
  ASSERT_TRUE(static_cast<bool>(wb.project_pivot(s, 0)));
  const pivot::PivotRect after = *table.projected_extent();
  EXPECT_EQ(after.row0, moved.row0);
  EXPECT_EQ(after.row1, moved.row1 - 1U);
  for (std::uint32_t c = moved.col0; c <= moved.col1; ++c) {
    EXPECT_FALSE(Facts(wb, wb.sheet(s), moved.row1, c).present) << A1(moved.row1, c);
  }
  for (std::uint32_t r = moved.row0; r < moved.row1; ++r) {
    for (std::uint32_t c = moved.col0; c <= moved.col1; ++c) {
      const auto it = excel.find({r - 2U, c});
      const CellFacts want = it == excel.end() ? CellFacts{} : it->second;
      const CellFacts got = Facts(wb, wb.sheet(s), r, c);
      EXPECT_EQ(got.value, want.value) << A1(r, c);
      EXPECT_EQ(got.format, want.format) << A1(r, c);
    }
  }
}

TEST(PivotProjectionFixture, PendingPivotsAreProjectedAndCleared) {
  std::unique_ptr<io::OoxmlReadResult> read = ReadFixture("pivot_layout/ja_c0_rows.xlsx");
  ASSERT_NE(read, nullptr);
  Workbook& wb = read->workbook;
  wb.set_excel_profile(FixtureProfile("pivot_layout/ja_c0_rows.xlsx"));
  const std::size_t s = PivotSheet(wb);
  pivot::PivotTable& table = *wb.sheet(s).mutable_pivot_tables()[0];
  const std::uint32_t row = table.anchor_row() + 1U;
  const std::uint32_t col = table.anchor_col();
  const CellFacts excel = Facts(wb, wb.sheet(s), row, col);
  ASSERT_TRUE(excel.present);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(s, row, col, "stale")));

  // Nothing pending: the stale cell stays.
  wb.project_pending_pivots();
  EXPECT_EQ(Facts(wb, wb.sheet(s), row, col).value, "s:stale");

  table.set_projection_pending(true);
  wb.project_pending_pivots();
  EXPECT_FALSE(table.projection_pending());
  EXPECT_EQ(Facts(wb, wb.sheet(s), row, col).value, excel.value);
  EXPECT_EQ(Facts(wb, wb.sheet(s), row, col).format, excel.format);
}

TEST(PivotProjectionFixture, OverlapRefusalWritesNothing) {
  std::unique_ptr<io::OoxmlReadResult> read = ReadFixture("pivot_layout/ja_c0_rows.xlsx");
  ASSERT_NE(read, nullptr);
  Workbook& wb = read->workbook;
  wb.set_excel_profile(FixtureProfile("pivot_layout/ja_c0_rows.xlsx"));
  const std::size_t s = PivotSheet(wb);
  pivot::PivotTable& table = *wb.sheet(s).mutable_pivot_tables()[0];
  // An array block where the report would land after a move.
  ASSERT_TRUE(static_cast<bool>(wb.set_array_formula(s, 41U, 21U, 2U, 1U, "={1;2}")));
  const auto before = Snapshot(wb, wb.sheet(s));
  const std::size_t merges = wb.sheet(s).merges().size();
  table.set_anchor(40U, 20U, table.span_rows(), table.span_cols());
  auto projected_or = wb.project_pivot(s, 0);
  ASSERT_FALSE(static_cast<bool>(projected_or));
  EXPECT_EQ(projected_or.error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_FALSE(table.projected_extent().has_value());

  const auto after = Snapshot(wb, wb.sheet(s));
  ASSERT_EQ(after.size(), before.size());
  for (const auto& [at, facts] : before) {
    const auto it = after.find(at);
    ASSERT_NE(it, after.end()) << A1(at.first, at.second);
    EXPECT_EQ(it->second.value, facts.value) << A1(at.first, at.second);
    EXPECT_EQ(it->second.format, facts.format) << A1(at.first, at.second);
  }
  EXPECT_EQ(wb.sheet(s).merges().size(), merges);
}

TEST(PivotProjectionFixture, ACacheSavedWithoutRecordsWritesNothing) {
  std::unique_ptr<io::OoxmlReadResult> read = ReadFixture("pivot_options/op_savedata_off.xlsx");
  ASSERT_NE(read, nullptr);
  Workbook& wb = read->workbook;
  const std::size_t s = PivotSheet(wb);
  const std::uint32_t row = wb.sheet(s).pivot_tables()[0]->anchor_row();
  const std::uint32_t col = wb.sheet(s).pivot_tables()[0]->anchor_col();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(s, row, col, "kept")));
  ASSERT_TRUE(static_cast<bool>(wb.project_pivot(s, 0)));
  EXPECT_EQ(Facts(wb, wb.sheet(s), row, col).value, "s:kept");
  EXPECT_FALSE(wb.sheet(s).pivot_tables()[0]->projected_extent().has_value());
}

TEST(PivotProjectionFixture, ClearingTheProjectionErasesItsCellsAndMerges) {
  std::unique_ptr<io::OoxmlReadResult> read = ReadFixture("pivot_options/op_merge_tabular.xlsx");
  ASSERT_NE(read, nullptr);
  Workbook& wb = read->workbook;
  wb.set_excel_profile(FixtureProfile("pivot_options/op_merge_tabular.xlsx"));
  const std::size_t s = PivotSheet(wb);
  ASSERT_TRUE(static_cast<bool>(wb.project_pivot(s, 0)));
  const pivot::PivotRect e = *wb.sheet(s).pivot_tables()[0]->projected_extent();
  ASSERT_TRUE(static_cast<bool>(wb.clear_pivot_projection(s, 0)));
  EXPECT_TRUE(wb.sheet(s).merges().empty());
  for (std::uint32_t r = e.row0; r <= e.row1; ++r) {
    for (std::uint32_t c = e.col0; c <= e.col1; ++c) {
      EXPECT_EQ(wb.sheet(s).cell_at(r, c), nullptr) << A1(r, c);
    }
  }
  EXPECT_FALSE(wb.sheet(s).pivot_tables()[0]->projected_extent().has_value());
}

}  // namespace
}  // namespace formulon
