//
// `<rowItems>` / `<colItems>` generated for every Excel-saved pivot of the
// fixture corpus, against the elements Excel wrote.
//
// Each pivot of `tests/fixtures/excel/**/*.xlsx` (and of the pivot_refresh
// oracle cases' before/after books, when present) is evaluated and projected
// under the profile it was saved in, and its line items are serialized as an
// edited table's would be on save. Each axis then lands in one bucket:
//
//   * encoded: byte-equal to the element read from the file;
//   * refused: the projection declines to encode the axis' shape;
//   * divergent: the projected cells differ from the cells Excel saved, so
//     the lines cannot be expected to agree either;
//   * unmapped: a line names an item no written `<items>` entry can stand
//     for unambiguously, so the table saves no items (both axes).
//
// The refused, divergent and unmapped sets are pinned exactly and the encoded
// count has a floor; anything else -- an unequal element over a matching
// projection -- fails. The observed
// buckets are printed in pinnable form on every run.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <filesystem>
#include <memory>
#include <optional>
#include <set>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/pivot_cache_writer.h"
#include "io/pivot_line_items_writer.h"
#include "io/pivot_table_writer.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

// Pinned buckets. A key whose file is absent from this checkout (the
// pivot_refresh cases are optional) is left out of the comparison.

/// Axes of the fixture tree (pivot_refresh cases excluded) encoded
/// byte-equal to the element Excel wrote: at least this many.
constexpr std::size_t kEncodedFloor = 659;

/// Axes whose shape the projection declines to encode, as "<file>#s<sheet>p<pivot>:<axis>".
const std::set<std::string>& PinnedRefused() {
  static const std::set<std::string> kRefused = {
      "pivot_line_shapes/M1a_values_above_two_rows_compact.xlsx#s0p0:rows",
      "pivot_line_shapes/M1b_values_above_two_rows_outline.xlsx#s0p0:rows",
      "pivot_line_shapes/M1c_values_above_two_rows_tabular.xlsx#s0p0:rows",
      "pivot_line_shapes/M2b_collapsed_outer_mid_values_below.xlsx#s0p0:rows",
      "pivot_line_shapes/M2c_collapsed_outer_mid_values_below_outline.xlsx#s0p0:rows",
      "pivot_line_shapes/M3_collapsed_inner_tabular.xlsx#s0p0:rows",
      "pivot_line_shapes/M4a_custom_subtotals_outer_values_below.xlsx#s0p0:rows",
      "pivot_line_shapes/M4b_custom_subtotals_outer_values_below_tabular.xlsx#s0p0:rows",
      "pivot_line_shapes/M5_custom_subtotals_col_field.xlsx#s0p0:cols",
      "pivot_line_shapes/M5b_custom_subtotals_col_field_with_rows.xlsx#s0p0:cols",
      "pivot_line_shapes/M6a_page_field_only.xlsx#s0p0:cols",
      "pivot_line_shapes/M6a_page_field_only.xlsx#s0p0:rows",
      "pivot_line_shapes/M6b_completely_empty.xlsx#s0p0:cols",
      "pivot_line_shapes/M6b_completely_empty.xlsx#s0p0:rows",
      "pivot_line_shapes/M7a_outer_outline_inner_compact.xlsx#s0p0:rows",
      "pivot_line_shapes/M7b_outer_tabular_inner_outline.xlsx#s0p0:rows",
      "pivot_options/collapse_item_outer_column.xlsx#s0p0:cols",
      "pivot_options/op_emptycol_show_all.xlsx#s0p0:cols",
      "pivot_options/op_emptyrow_show_all.xlsx#s0p0:rows",
      "pivot_options/op_showall_only.xlsx#s0p0:rows",
      "pivot_refresh/hdr_dup/after.xlsx#s0p0:cols",
      "pivot_refresh/hdr_dup/after.xlsx#s0p0:rows",
      "pivot_refresh/hdr_dup_suffix_collision/after.xlsx#s0p0:cols",
      "pivot_refresh/hdr_dup_suffix_collision/after.xlsx#s0p0:rows",
      "pivot_refresh/hdr_kana_dup/after.xlsx#s0p0:cols",
      "pivot_refresh/hdr_kana_dup/after.xlsx#s0p0:rows",
      "pivot_refresh/hdr_width_dup/after.xlsx#s0p0:cols",
      "pivot_refresh/hdr_width_dup/after.xlsx#s0p0:rows",
      "pivot_refresh/type_basic_inline/after.xlsx#s0p0:cols",
      "pivot_refresh/type_basic_inline/after.xlsx#s0p0:rows",
      "pivot_refresh/type_basic_inline/before.xlsx#s0p0:cols",
      "pivot_refresh/type_basic_inline/before.xlsx#s0p0:rows",
      "pivot_refresh/type_date_number_bytes/after.xlsx#s0p0:cols",
      "pivot_refresh/type_date_number_bytes/after.xlsx#s0p0:rows",
      "pivot_refresh/type_date_number_bytes/before.xlsx#s0p0:cols",
      "pivot_refresh/type_date_number_bytes/before.xlsx#s0p0:rows",
      "pivot_refresh/type_dates_inline/after.xlsx#s0p0:cols",
      "pivot_refresh/type_dates_inline/after.xlsx#s0p0:rows",
      "pivot_refresh/type_dates_inline/before.xlsx#s0p0:cols",
      "pivot_refresh/type_dates_inline/before.xlsx#s0p0:rows",
      "xlsb_pivot_twins/B04_rowgrand_off.xlsx#s0p0:cols",
      "xlsb_pivot_twins/C01_subtotal_bottom.xlsx#s0p0:rows",
      "xlsb_pivot_twins/C04_show_all.xlsx#s0p0:rows",
  };
  return kRefused;
}

/// Pivots whose projection differs from the report Excel saved, as "<file>#s<sheet>p<pivot>".
const std::set<std::string>& PinnedDivergent() {
  static const std::set<std::string> kDivergent = {
      "pivot_options/op_savedata_off.xlsx#s0p0",
      "pivot_refresh/g2_numgroup_auto_widen/after.xlsx#s3p0",
      "pivot_refresh/g2_numgroup_auto_widen/before.xlsx#s3p0",
      "pivot_refresh/g3_numgroup_fixed_widen/after.xlsx#s3p0",
      "pivot_refresh/g3_numgroup_fixed_widen/before.xlsx#s3p0",
      "pivot_refresh/g4_numgroup_auto_control/after.xlsx#s3p0",
      "pivot_refresh/g4_numgroup_auto_control/before.xlsx#s3p0",
      "xlsb_pivot_twins/C07_sort_by_data.xlsx#s0p0",
  };
  return kDivergent;
}

/// Pivots with a line no written item stands for unambiguously, as "<file>#s<sheet>p<pivot>".
const std::set<std::string>& PinnedUnmapped() {
  static const std::set<std::string> kUnmapped = {
      "pivot_refresh/type_basic_rows/after.xlsx#s0p0",
      "pivot_refresh/type_dates_rows/after.xlsx#s0p0",
  };
  return kUnmapped;
}

/// Every .xlsx under the Excel fixture tree, then the pivot_refresh cases'
/// before/after books, keyed by a path relative to their root.
std::vector<std::pair<std::string, std::filesystem::path>> CorpusFiles() {
  std::set<std::pair<std::string, std::filesystem::path>> files;
  const std::filesystem::path excel = std::filesystem::path(FORMULON_FIXTURES_DIR) / "excel";
  std::error_code ec;
  for (std::filesystem::recursive_directory_iterator it(excel, ec), end; !ec && it != end; it.increment(ec)) {
    if (it->is_regular_file() && it->path().extension() == ".xlsx") {
      files.emplace(std::filesystem::relative(it->path(), excel).generic_string(), it->path());
    }
  }
  EXPECT_FALSE(ec) << excel << ": " << ec.message();
  // Optional inputs: skipped when the oracle cases are absent.
  const std::filesystem::path refresh =
      std::filesystem::path(FORMULON_FIXTURES_DIR) / ".." / "oracle" / "cases_wb" / "pivot_refresh";
  std::error_code refresh_ec;
  for (std::filesystem::directory_iterator it(refresh, refresh_ec), end; !refresh_ec && it != end;
       it.increment(refresh_ec)) {
    for (const char* book : {"before.xlsx", "after.xlsx"}) {
      const std::filesystem::path path = it->path() / book;
      if (std::filesystem::is_regular_file(path)) {
        files.emplace("pivot_refresh/" + it->path().filename().string() + "/" + book, path);
      }
    }
  }
  return {files.begin(), files.end()};
}

/// The profile a book was saved under: English books carry `en_` at the start
/// of a path segment, the `win/` tree is Windows, everything else Mac ja-JP.
ExcelProfile FixtureProfile(const std::string& relative) {
  const char* id = "mac-365-ja_JP";
  if (relative.rfind("en_", 0) == 0 || relative.find("/en_") != std::string::npos) {
    id = "mac-365-en_US";
  } else if (relative.rfind("win/", 0) == 0) {
    id = "win-365-ja_JP";
  }
  ExcelProfile profile;
  EXPECT_TRUE(parse_excel_profile_id(id, &profile)) << id;
  return profile;
}

std::vector<std::uint8_t> ReadFileBytes(const std::filesystem::path& path) {
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.string().c_str(), "rb");
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

bool SameValue(const Value& a, const Value& b) {
  if (a.is_number() && b.is_number()) {
    const double x = a.as_number();
    const double y = b.as_number();
    return x == y || std::fabs(x - y) <= 1e-9 * std::fmax(std::fabs(x), std::fabs(y));
  }
  if (a.is_text() && b.is_text()) {
    return a.as_text() == b.as_text();
  }
  if (a.is_boolean() && b.is_boolean()) {
    return a.as_boolean() == b.as_boolean();
  }
  if (a.is_error() && b.is_error()) {
    return a.as_error() == b.as_error();
  }
  return a.is_blank() && b.is_blank();
}

/// The first projected cell whose value differs from the sheet's, or empty.
std::string FirstDivergence(const pivot::PivotCells& cells, const Sheet& sheet) {
  for (const pivot::PivotCell& cell : cells.cells) {
    const Value got = cell.item_value.has_value() ? *cell.item_value : cell.value;
    const Cell* saved = sheet.cell_at(cell.row, cell.col);
    const Value want = saved != nullptr ? saved->cached_value : Value::blank();
    if (!SameValue(got, want)) {
      std::ostringstream os;
      os << "row " << cell.row << " col " << cell.col;
      return os.str();
    }
  }
  return {};
}

struct Buckets {
  std::size_t pivots = 0;
  std::size_t encoded = 0;
  std::size_t encoded_fixtures = 0;  ///< `encoded` outside the optional pivot_refresh cases.
  std::set<std::string> files;       ///< Relative paths read, to scope the pinned sets.
  std::set<std::string> refused;
  std::set<std::string> divergent;
  std::set<std::string> unmapped;
  std::vector<std::string> failures;
};

/// Sorts both axes of one pivot into `buckets`.
void Classify(const std::string& key, const Workbook& wb, const Sheet& sheet, const pivot::PivotTable& table,
              const ExcelProfile& profile, Buckets& buckets) {
  const pivot::PivotCache* cache = wb.find_pivot_cache(table.pivot_cache_id());
  if (cache == nullptr) {
    buckets.divergent.insert(key);
    return;
  }
  // The write plan and the item matching collate text, as on save.
  const eval::ScopedEvalProfile scope(profile);
  std::vector<const pivot::PivotTable*> users;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const std::unique_ptr<pivot::PivotTable>& other : wb.sheet(s).pivot_tables()) {
      if (other != nullptr && other->pivot_cache_id() == table.pivot_cache_id()) {
        users.push_back(other.get());
      }
    }
  }
  const io::PivotCacheWritePlan plan = io::build_pivot_cache_write_plan(*cache, users, wb.date1904());
  const pivot::PivotLayoutOptions options = pivot::pivot_layout_options_for(profile);
  auto result_or =
      pivot::evaluate(table, *cache, options, pivot::PivotFilterEnv{wb.pinned_now(), wb.date1904(), &wb.styles()});
  if (!result_or) {
    buckets.divergent.insert(key);
    return;
  }
  pivot::PivotLineItems lines;
  auto cells_or = pivot::layout(table, result_or.value(), options, &lines);
  if (!cells_or) {
    buckets.divergent.insert(key);
    return;
  }
  // Serialize whichever axes were not refused; a refused one writes an empty list.
  pivot::PivotLineItems encodable = lines;
  encodable.row_refusal.reset();
  encodable.col_refusal.reset();
  const std::optional<io::PivotItemsXml> xml = io::write_pivot_line_items(table, *cache, &plan, encodable);
  const bool diverges = !FirstDivergence(cells_or.value(), sheet).empty();
  if (!xml.has_value()) {
    buckets.unmapped.insert(key);
  }
  struct Axis {
    const char* name;
    bool refused;
    const std::string* read;
    const std::string* written;
  };
  const Axis axes[] = {
      {"rows", lines.row_refusal.has_value(), &table.raw_passthrough_after_row_fields(),
       xml ? &xml->row_items : nullptr},
      {"cols", lines.col_refusal.has_value(), &table.raw_col_items(), xml ? &xml->col_items : nullptr},
  };
  for (const Axis& axis : axes) {
    const std::string axis_key = key + ":" + axis.name;
    if (axis.refused) {
      buckets.refused.insert(axis_key);
    } else if (axis.written == nullptr) {
      continue;
    } else if (*axis.written == *axis.read) {
      ++buckets.encoded;
      buckets.encoded_fixtures += key.rfind("pivot_refresh/", 0) == 0 ? 0U : 1U;
    } else if (diverges) {
      buckets.divergent.insert(key);
    } else {
      buckets.failures.push_back(axis_key + ":\n    excel:  " + *axis.read + "\n    engine: " + *axis.written);
    }
  }
}

/// The pinned keys whose file this run read.
std::set<std::string> PresentIn(const std::set<std::string>& pinned, const std::set<std::string>& files) {
  std::set<std::string> out;
  for (const std::string& key : pinned) {
    if (files.count(key.substr(0, key.find('#'))) != 0U) {
      out.insert(key);
    }
  }
  return out;
}

std::string Pinnable(const std::set<std::string>& keys) {
  std::string out;
  for (const std::string& key : keys) {
    out += "      \"" + key + "\",\n";
  }
  return out;
}

TEST(PivotLineItemsCorpus, EveryExcelPivotEncodesRefusesOrDiverges) {
  Buckets buckets;
  for (const auto& [relative, path] : CorpusFiles()) {
    const std::vector<std::uint8_t> bytes = ReadFileBytes(path);
    if (bytes.empty()) {
      ADD_FAILURE() << "could not read " << path;
      continue;
    }
    auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
    if (!read_or) {
      // A book that does not load is another suite's concern; only pivots count here.
      continue;
    }
    buckets.files.insert(relative);
    Workbook& wb = read_or.value().workbook;
    const ExcelProfile profile = FixtureProfile(relative);
    wb.set_excel_profile(profile);
    for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
      const Sheet& sheet = wb.sheet(s);
      for (std::size_t p = 0; p < sheet.pivot_tables().size(); ++p) {
        const pivot::PivotTable* table = sheet.pivot_tables()[p].get();
        if (table == nullptr) {
          continue;
        }
        ++buckets.pivots;
        Classify(relative + "#s" + std::to_string(s) + "p" + std::to_string(p), wb, sheet, *table, profile, buckets);
      }
    }
  }
  std::printf("[ line items ] %zu pivots, %zu axes encoded (%zu outside pivot_refresh)\n", buckets.pivots,
              buckets.encoded, buckets.encoded_fixtures);
  std::printf("[ line items ] refused:\n%s", Pinnable(buckets.refused).c_str());
  std::printf("[ line items ] divergent:\n%s", Pinnable(buckets.divergent).c_str());
  std::printf("[ line items ] unmapped:\n%s", Pinnable(buckets.unmapped).c_str());
  RecordProperty("pivots", static_cast<int>(buckets.pivots));
  RecordProperty("encoded", static_cast<int>(buckets.encoded));

  EXPECT_GT(buckets.pivots, 0U);
  EXPECT_GE(buckets.encoded_fixtures, kEncodedFloor);
  EXPECT_EQ(buckets.refused, PresentIn(PinnedRefused(), buckets.files)) << "observed refused set:\n"
                                                                        << Pinnable(buckets.refused);
  EXPECT_EQ(buckets.divergent, PresentIn(PinnedDivergent(), buckets.files)) << "observed divergent set:\n"
                                                                            << Pinnable(buckets.divergent);
  EXPECT_EQ(buckets.unmapped, PresentIn(PinnedUnmapped(), buckets.files)) << "observed unmapped set:\n"
                                                                          << Pinnable(buckets.unmapped);
  std::string failures;
  for (const std::string& failure : buckets.failures) {
    failures += "  " + failure + "\n";
  }
  EXPECT_TRUE(buckets.failures.empty()) << buckets.failures.size() << " unexpected axes:\n" << failures;
}

}  // namespace
}  // namespace formulon
