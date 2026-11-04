//
// Workbooks Mac Excel 365 saved after setting a pivot report option
// (grand-total caption, blank rows, repeated labels, subtotals, collapse
// state, item captions, autofit, date grouping) through its own UI
// automation.
//
// Load -> save -> reload: each option the saved definition spells must come
// back with the same attributes; `rowItems`, `colItems`, `location@ref`,
// `cacheId` and `refreshOnLoad` are left out because a save may rebuild them.
//
// Layout: the pivot evaluated and projected under the ja-JP profile the files
// were saved in must draw the text grid Excel drew.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <initializer_list>
#include <map>
#include <memory>
#include <optional>
#include <ostream>
#include <sstream>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/pivot_attr_names.h"
#include "miniz.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_table.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::string FixturePath(const char* file) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_options/" + file;
}

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

/// Every package entry whose name starts with `prefix`, keyed by entry name.
std::map<std::string, std::string> EntriesWithPrefix(const std::vector<std::uint8_t>& archive,
                                                     std::string_view prefix) {
  std::map<std::string, std::string> out;
  mz_zip_archive reader{};
  if (mz_zip_reader_init_mem(&reader, archive.data(), archive.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_reader_init_mem failed";
    return out;
  }
  for (mz_uint i = 0; i < mz_zip_reader_get_num_files(&reader); ++i) {
    char name[512] = {};
    if (mz_zip_reader_get_filename(&reader, i, name, sizeof(name)) == 0 ||
        std::string_view(name).substr(0, prefix.size()) != prefix) {
      continue;
    }
    std::size_t size = 0;
    void* body = mz_zip_reader_extract_to_heap(&reader, i, &size, 0);
    if (body != nullptr) {
      out.emplace(name, std::string(static_cast<const char*>(body), size));
      mz_free(body);
    }
  }
  mz_zip_reader_end(&reader);
  return out;
}

/// The parsed parts under `prefix` that are not relationship parts, in entry
/// name order.
std::vector<std::unique_ptr<pugi::xml_document>> ParseParts(const std::vector<std::uint8_t>& archive,
                                                            std::string_view prefix) {
  std::vector<std::unique_ptr<pugi::xml_document>> docs;
  for (const auto& [name, body] : EntriesWithPrefix(archive, prefix)) {
    if (name.find("/_rels/") != std::string::npos) {
      continue;
    }
    auto doc = std::make_unique<pugi::xml_document>();
    EXPECT_TRUE(doc->load_buffer(body.data(), body.size())) << name;
    docs.push_back(std::move(doc));
  }
  return docs;
}

/// `name="value"` for each listed attribute, `name=(absent)` when missing.
std::string Attrs(const pugi::xml_node& node, std::initializer_list<const char*> names) {
  std::string out;
  for (const char* name : names) {
    const pugi::xml_attribute attr = node.attribute(name);
    out.append(name).append(attr ? "=\"" + std::string(attr.value()) + "\" " : "=(absent) ");
  }
  return out;
}

/// Every attribute of `node`, sorted by name.
std::string AllAttrs(const pugi::xml_node& node) {
  std::map<std::string, std::string> sorted;
  for (const pugi::xml_attribute attr : node.attributes()) {
    sorted.emplace(attr.name(), attr.value());
  }
  std::string out;
  for (const auto& [name, value] : sorted) {
    out.append(name).append("=\"").append(value).append("\" ");
  }
  return out;
}

/// The option-bearing state of one `<pivotTableDefinition>`, one line per
/// fact, so a mismatch reports exactly what moved.
std::string TableOptions(const pugi::xml_node& root) {
  std::string out = "root " + Attrs(root, {"grandTotalCaption", "useAutoFormatting"}) + "\n";
  std::size_t index = 0;
  for (const pugi::xml_node field : root.child("pivotFields").children("pivotField")) {
    out.append("field ").append(std::to_string(index++)).append(" ");
    out.append(Attrs(field, {"insertBlankRow", "defaultSubtotal"}));
    std::string fill_down = "(absent)";
    for (const pugi::xml_node ext : field.child("extLst").children("ext")) {
      if (strings::case_insensitive_eq(ext.attribute("uri").value(), io::kPivotFieldExtUri)) {
        fill_down = ext.child("x14:pivotField").attribute("fillDownLabels").value();
      }
    }
    out.append("fillDownLabels=").append(fill_down).append("\n");
    for (const pugi::xml_node item : field.child("items").children("item")) {
      out.append("  item ").append(Attrs(item, {"t", "x", "h", "n", "sd"})).append("\n");
    }
  }
  return out;
}

/// The cache fields of one `<pivotCacheDefinition>` with their grouping.
std::string CacheOptions(const pugi::xml_node& root) {
  std::string out;
  for (const pugi::xml_node field : root.child("cacheFields").children("cacheField")) {
    out.append("cacheField ").append(Attrs(field, {"name", "numFmtId", "databaseField"})).append("\n");
    const pugi::xml_node group = field.child("fieldGroup");
    if (!group) {
      continue;
    }
    out.append("  fieldGroup ").append(Attrs(group, {"par", "base"})).append("\n");
    if (const pugi::xml_node range = group.child("rangePr"); range) {
      out.append("  rangePr ").append(AllAttrs(range)).append("\n");
    }
    for (const pugi::xml_node x : group.child("discretePr").children("x")) {
      out.append("  discrete ").append(x.attribute("v").value()).append("\n");
    }
    for (const pugi::xml_node item : group.child("groupItems").children()) {
      out.append("  groupItem ").append(item.name()).append(" ").append(item.attribute("v").value()).append("\n");
    }
  }
  return out;
}

/// Table name and `location@ref`: tables of one workbook can share a name.
std::string TableKey(const pugi::xml_node& root) {
  return std::string(root.attribute("name").value()) + " @ " + root.child("location").attribute("ref").value();
}

std::vector<std::uint8_t> LoadAndSave(const std::vector<std::uint8_t>& bytes) {
  auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;
  if (!read_or) {
    return {};
  }
  auto saved_or = io::write_ooxml(read_or.value().workbook);
  EXPECT_TRUE(static_cast<bool>(saved_or)) << saved_or.error().message;
  return saved_or ? std::move(saved_or.value()) : std::vector<std::uint8_t>{};
}

class PivotOptionsFixture : public ::testing::TestWithParam<const char*> {};

TEST_P(PivotOptionsFixture, RoundTripKeepsEveryOptionExcelWrote) {
  const std::vector<std::uint8_t> original = ReadFileBytes(FixturePath(GetParam()));
  ASSERT_FALSE(original.empty()) << "could not read " << GetParam();
  const std::vector<std::uint8_t> saved = LoadAndSave(original);
  ASSERT_FALSE(saved.empty());
  auto reread_or = io::read_ooxml(io::ByteSpan{saved.data(), saved.size()});
  ASSERT_TRUE(static_cast<bool>(reread_or)) << reread_or.error().message;

  std::map<std::string, std::string> saved_tables;
  for (const auto& doc : ParseParts(saved, "xl/pivotTables/")) {
    const pugi::xml_node root = doc->child("pivotTableDefinition");
    saved_tables.emplace(TableKey(root), TableOptions(root));
  }
  const auto original_tables = ParseParts(original, "xl/pivotTables/");
  ASSERT_FALSE(original_tables.empty());
  ASSERT_EQ(saved_tables.size(), original_tables.size());
  for (const auto& doc : original_tables) {
    const pugi::xml_node root = doc->child("pivotTableDefinition");
    const auto it = saved_tables.find(TableKey(root));
    ASSERT_NE(it, saved_tables.end()) << "no saved table " << TableKey(root);
    EXPECT_EQ(it->second, TableOptions(root)) << TableKey(root);
  }

  const auto original_caches = ParseParts(original, "xl/pivotCache/pivotCacheDefinition");
  const auto saved_caches = ParseParts(saved, "xl/pivotCache/pivotCacheDefinition");
  ASSERT_EQ(saved_caches.size(), original_caches.size());
  for (std::size_t i = 0; i < original_caches.size(); ++i) {
    EXPECT_EQ(CacheOptions(saved_caches[i]->child("pivotCacheDefinition")),
              CacheOptions(original_caches[i]->child("pivotCacheDefinition")))
        << "cache " << i;
  }
}

std::string FixtureName(const ::testing::TestParamInfo<const char*>& info) {
  std::string name;
  for (const char* p = info.param; *p != '\0' && *p != '.'; ++p) {
    name.push_back(*p);
  }
  return name;
}

INSTANTIATE_TEST_SUITE_P(
    ExcelSaved, PivotOptionsFixture,
    ::testing::Values("grand_total_caption_tabular.xlsx", "grand_total_caption_values_on_rows.xlsx",
                      "blank_row_outer_tabular.xlsx", "blank_row_outer_outline.xlsx", "blank_row_single_compact.xlsx",
                      "blank_row_both_tabular.xlsx", "repeat_labels_tabular.xlsx", "repeat_labels_outline.xlsx",
                      "repeat_labels_compact.xlsx", "subtotal_off_outer_tabular.xlsx", "collapse_item_tabular.xlsx",
                      "collapse_item_outline.xlsx", "collapse_item_compact.xlsx", "collapse_field_outline.xlsx",
                      "expand_field_outline.xlsx", "show_details_leaf.xlsx", "show_details_subtotal_and_grand.xlsx",
                      "show_details_twice.xlsx", "item_caption.xlsx", "auto_formatting_off.xlsx",
                      "auto_formatting_on.xlsx", "date_field_ungrouped.xlsx", "date_group_months.xlsx",
                      "date_group_months_years.xlsx", "date_group_quarters_months_years.xlsx", "date_group_days.xlsx",
                      "date_group_days_interval7.xlsx", "date_group_years.xlsx", "date_group_quarters.xlsx",
                      "date_group_ungrouped.xlsx"),
    FixtureName);

// ---------------------------------------------------------------------------
// An edited table hands its report back to Excel to rebuild
// ---------------------------------------------------------------------------

/// The one pivot table of the tabular grand-total fixture, read for editing.
struct LoadedTable {
  std::unique_ptr<io::OoxmlReadResult> read;
  pivot::PivotTable* table = nullptr;
  pivot::PivotCache* cache = nullptr;
};

LoadedTable LoadForEdit() {
  LoadedTable out;
  const std::vector<std::uint8_t> bytes = ReadFileBytes(FixturePath("grand_total_caption_tabular.xlsx"));
  auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;
  if (!read_or) {
    return out;
  }
  out.read = std::make_unique<io::OoxmlReadResult>(std::move(read_or.value()));
  Workbook& wb = out.read->workbook;
  for (std::size_t i = 0; i < wb.sheet_count() && out.table == nullptr; ++i) {
    if (!wb.sheet(i).pivot_tables().empty()) {
      out.table = wb.sheet(i).mutable_pivot_tables().front().get();
    }
  }
  if (out.table != nullptr) {
    for (const std::unique_ptr<pivot::PivotCache>& cache : wb.mutable_pivot_caches()) {
      if (cache != nullptr && cache->cache_id() == out.table->pivot_cache_id()) {
        out.cache = cache.get();
      }
    }
  }
  return out;
}

/// The single pivot table part and cache definition part of a saved package.
std::pair<std::string, std::string> SavedPivotParts(const Workbook& wb) {
  auto saved_or = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved_or)) << saved_or.error().message;
  if (!saved_or) {
    return {};
  }
  std::pair<std::string, std::string> parts;
  for (const auto& [name, body] : EntriesWithPrefix(saved_or.value(), "xl/pivot")) {
    if (name.rfind("xl/pivotTables/pivotTable", 0) == 0) {
      parts.first = body;
    } else if (name.rfind("xl/pivotCache/pivotCacheDefinition", 0) == 0) {
      parts.second = body;
    }
  }
  return parts;
}

TEST(PivotOptionsFixture, RoundTripUntouchedTableKeepsItsRenderedItems) {
  LoadedTable loaded = LoadForEdit();
  ASSERT_NE(loaded.table, nullptr);
  const auto [table_xml, cache_xml] = SavedPivotParts(loaded.read->workbook);
  EXPECT_NE(table_xml.find("<rowItems"), std::string::npos) << table_xml;
  EXPECT_NE(table_xml.find("<colItems"), std::string::npos) << table_xml;
  EXPECT_EQ(cache_xml.find("refreshOnLoad"), std::string::npos) << cache_xml;
}

TEST(PivotOptionsFixture, RoundTripEditedTableDropsItsItemsAndRefreshesOnLoad) {
  LoadedTable loaded = LoadForEdit();
  ASSERT_NE(loaded.table, nullptr);
  loaded.table->set_grand_total_caption(std::string("ALL"));
  loaded.table->mark_layout_dirty();
  const auto [table_xml, cache_xml] = SavedPivotParts(loaded.read->workbook);
  EXPECT_EQ(table_xml.find("<rowItems"), std::string::npos) << table_xml;
  EXPECT_EQ(table_xml.find("<colItems"), std::string::npos) << table_xml;
  EXPECT_NE(table_xml.find(" grandTotalCaption=\"ALL\""), std::string::npos) << table_xml;
  EXPECT_NE(cache_xml.find(" refreshOnLoad=\"1\""), std::string::npos) << cache_xml;
}

TEST(PivotOptionsFixture, RoundTripEditedTableOverAnOutsideSourceKeepsItsItems) {
  LoadedTable loaded = LoadForEdit();
  ASSERT_NE(loaded.table, nullptr);
  ASSERT_NE(loaded.cache, nullptr);
  ASSERT_TRUE(loaded.cache->worksheet_source().present);
  loaded.cache->mutable_worksheet_source().sheet = "No Such Sheet";
  loaded.table->mark_layout_dirty();
  const auto [table_xml, cache_xml] = SavedPivotParts(loaded.read->workbook);
  // Excel would not rebuild the report, so the items read with it stay.
  EXPECT_NE(table_xml.find("<rowItems"), std::string::npos) << table_xml;
  EXPECT_NE(table_xml.find("<colItems"), std::string::npos) << table_xml;
  EXPECT_EQ(cache_xml.find("refreshOnLoad"), std::string::npos) << cache_xml;
}

// ---------------------------------------------------------------------------
// The report each option makes Excel draw
// ---------------------------------------------------------------------------

std::string CellText(const Value& v) {
  if (v.is_text()) {
    return std::string(v.as_text());
  }
  if (v.is_number()) {
    std::ostringstream os;
    os << v.as_number();
    return os.str();
  }
  return v.is_blank() ? std::string() : std::string("?");
}

/// The pivot named `table_name` on `sheet_name`, or the workbook's only pivot
/// when both are null, projected under the ja-JP profile.
std::optional<pivot::PivotCells> ProjectFixture(const char* file, const char* sheet_name = nullptr,
                                                const char* table_name = nullptr) {
  const std::vector<std::uint8_t> bytes = ReadFileBytes(FixturePath(file));
  EXPECT_FALSE(bytes.empty()) << "could not read " << file;
  auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << (read_or ? "" : read_or.error().message);
  if (!read_or) {
    return std::nullopt;
  }
  const Workbook& wb = read_or.value().workbook;
  const pivot::PivotTable* table = nullptr;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (sheet_name != nullptr && wb.sheet(i).name() != sheet_name) {
      continue;
    }
    for (const std::unique_ptr<pivot::PivotTable>& candidate : wb.sheet(i).pivot_tables()) {
      if (candidate != nullptr && (table_name == nullptr || candidate->name() == table_name)) {
        EXPECT_EQ(table, nullptr) << file << ": more than one pivot matches";
        table = candidate.get();
      }
    }
  }
  EXPECT_NE(table, nullptr) << file;
  const pivot::PivotCache* cache = table != nullptr ? wb.find_pivot_cache(table->pivot_cache_id()) : nullptr;
  EXPECT_NE(cache, nullptr) << file;
  if (cache == nullptr) {
    return std::nullopt;
  }
  const pivot::PivotLayoutOptions options = pivot::pivot_layout_options_for(mac_365_ja_jp_profile());
  auto result_or = pivot::evaluate(*table, *cache, options, pivot::PivotFilterEnv{std::nullopt, wb.date1904()});
  EXPECT_TRUE(static_cast<bool>(result_or)) << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return std::nullopt;
  }
  auto cells_or = pivot::layout(*table, result_or.value(), options);
  EXPECT_TRUE(static_cast<bool>(cells_or)) << (cells_or ? "" : cells_or.error().message);
  if (!cells_or) {
    return std::nullopt;
  }
  return std::move(cells_or.value());
}

/// The projected grid, one `|`-joined line per row.
std::vector<std::string> Grid(const pivot::PivotCells& cells) {
  std::vector<std::vector<std::string>> grid(cells.rows, std::vector<std::string>(cells.cols));
  for (const pivot::PivotCell& cell : cells.cells) {
    const std::uint32_t r = cell.row - cells.top;
    const std::uint32_t c = cell.col - cells.left;
    if (r < cells.rows && c < cells.cols) {
      grid[r][c] = CellText(cell.value);
    }
  }
  std::vector<std::string> lines;
  for (const std::vector<std::string>& row : grid) {
    std::string line;
    for (std::size_t c = 0; c < row.size(); ++c) {
      line += (c == 0 ? "" : "|") + row[c];
    }
    lines.push_back(std::move(line));
  }
  return lines;
}

// " 集計" / "総計" / "行ラベル" / "全体の ".
#define SUBTOTAL " \xE9\x9B\x86\xE8\xA8\x88"
#define TOTAL "\xE7\xB7\x8F\xE8\xA8\x88"
#define ROW_LABELS "\xE8\xA1\x8C\xE3\x83\xA9\xE3\x83\x99\xE3\x83\xAB"
#define ALL_OF "\xE5\x85\xA8\xE4\xBD\x93\xE3\x81\xAE "

/// One grid line from its cells.
std::string Line(std::initializer_list<const char*> cells) {
  std::string line;
  bool first = true;
  for (const char* cell : cells) {
    line += (first ? "" : "|") + std::string(cell);
    first = false;
  }
  return line;
}

#define COL_LABELS "\xE5\x88\x97\xE3\x83\xA9\xE3\x83\x99\xE3\x83\xAB"

/// The label rows of the two-column-field report: Region, Product and Values
/// over the data row, under the "Column Labels" row.
std::vector<std::string> ColumnReport(std::string region_row, std::string product_row, std::string values_row,
                                      std::string data_row) {
  const std::size_t cols = static_cast<std::size_t>(std::count(data_row.begin(), data_row.end(), '|')) + 1;
  return {COL_LABELS + std::string(cols - 1, '|'), std::move(region_row), std::move(product_row), std::move(values_row),
          std::move(data_row)};
}

struct GridCase {
  const char* file;
  std::vector<std::string> grid;
  const char* sheet = nullptr;  ///< With `table`, picks one pivot of several.
  const char* table = nullptr;
};

constexpr const char* kDateSheet = "Sheet5";
constexpr const char* kDateTable =
    "\xE3\x83\x94\xE3\x83\x9C\xE3\x83\x83\xE3\x83\x88\xE3\x83\x86\xE3\x83\xBC\xE3\x83\x96\xE3\x83\xAB"
    "4";  // ピボットテーブル4

void PrintTo(const GridCase& c, std::ostream* os) {
  *os << c.file;
}

class PivotOptionsFixtureGrid : public ::testing::TestWithParam<GridCase> {};

TEST_P(PivotOptionsFixtureGrid, LayoutDrawsTheGridExcelDrew) {
  const std::optional<pivot::PivotCells> cells = ProjectFixture(GetParam().file, GetParam().sheet, GetParam().table);
  ASSERT_TRUE(cells.has_value());
  EXPECT_EQ(Grid(*cells), GetParam().grid);
}

const char* const kTabularHeader = "Region|Product|Sum of Amt|Sum of Qty";
const char* const kCompactHeader = ROW_LABELS "|Summe von Amt|Summe von Qty";

INSTANTIATE_TEST_SUITE_P(
    ExcelSaved, PivotOptionsFixtureGrid,
    ::testing::Values(
        GridCase{
            "grand_total_caption_tabular.xlsx",
            {kTabularHeader, "East|A|500|5", "East" SUBTOTAL "||500|5", "North|A|100|1", "|B|200|2",
             "North" SUBTOTAL "||300|3", "South|A|300|3", "|B|400|4", "South" SUBTOTAL "||700|7", "TOTAL||1500|15"}},
        GridCase{"grand_total_caption_values_on_rows.xlsx",
                 {ROW_LABELS "|",
                  "East|",
                  "A|",
                  "Sum of Amt|500",
                  "Sum of Qty|5",
                  "East Sum of Amt|500",
                  "East Sum of Qty|5",
                  "North|",
                  "A|",
                  "Sum of Amt|100",
                  "Sum of Qty|1",
                  "B|",
                  "Sum of Amt|200",
                  "Sum of Qty|2",
                  "North Sum of Amt|300",
                  "North Sum of Qty|3",
                  "South|",
                  "A|",
                  "Sum of Amt|300",
                  "Sum of Qty|3",
                  "B|",
                  "Sum of Amt|400",
                  "Sum of Qty|4",
                  "South Sum of Amt|700",
                  "South Sum of Qty|7",
                  ALL_OF "Sum of Amt|1500",
                  ALL_OF "Sum of Qty|15"}},
        GridCase{"blank_row_outer_tabular.xlsx",
                 {kTabularHeader, "East|A|500|5", "East" SUBTOTAL "||500|5", "|||", "North|A|100|1", "|B|200|2",
                  "North" SUBTOTAL "||300|3", "|||", "South|A|300|3", "|B|400|4", "South" SUBTOTAL "||700|7", "|||",
                  TOTAL "||1500|15"}},
        GridCase{"blank_row_outer_outline.xlsx",
                 {kTabularHeader, "East||500|5", "|A|500|5", "|||", "North||300|3", "|A|100|1", "|B|200|2", "|||",
                  "South||700|7", "|A|300|3", "|B|400|4", "|||", TOTAL "||1500|15"}},
        GridCase{"blank_row_single_compact.xlsx",
                 {kCompactHeader, "East|500|5", "North|300|3", "South|700|7", TOTAL "|1500|15"}},
        GridCase{"blank_row_both_tabular.xlsx",
                 {kTabularHeader, "East|A|500|5", "East" SUBTOTAL "||500|5", "|||", "North|A|100|1", "|B|200|2",
                  "North" SUBTOTAL "||300|3", "|||", "South|A|300|3", "|B|400|4", "South" SUBTOTAL "||700|7", "|||",
                  TOTAL "||1500|15"}},
        GridCase{"repeat_labels_tabular.xlsx",
                 {kTabularHeader, "East|A|500|5", "East" SUBTOTAL "||500|5", "North|A|100|1", "North|B|200|2",
                  "North" SUBTOTAL "||300|3", "South|A|300|3", "South|B|400|4", "South" SUBTOTAL "||700|7",
                  TOTAL "||1500|15"}},
        GridCase{"repeat_labels_outline.xlsx",
                 {kTabularHeader, "East||500|5", "East|A|500|5", "North||300|3", "North|A|100|1", "North|B|200|2",
                  "South||700|7", "South|A|300|3", "South|B|400|4", TOTAL "||1500|15"}},
        GridCase{"repeat_labels_compact.xlsx",
                 {kCompactHeader, "East|500|5", "North|300|3", "South|700|7", TOTAL "|1500|15"}},
        GridCase{"subtotal_off_outer_tabular.xlsx",
                 {kTabularHeader, "East|A|500|5", "North|A|100|1", "|B|200|2", "South|A|300|3", "|B|400|4",
                  TOTAL "||1500|15"}},
        GridCase{"collapse_item_tabular.xlsx",
                 {kTabularHeader, "East||500|5", "North|A|100|1", "|B|200|2", "North" SUBTOTAL "||300|3",
                  "South|A|300|3", "|B|400|4", "South" SUBTOTAL "||700|7", TOTAL "||1500|15"}},
        GridCase{"collapse_item_outline.xlsx",
                 {kTabularHeader, "East||500|5", "North||300|3", "|A|100|1", "|B|200|2", "South||700|7", "|A|300|3",
                  "|B|400|4", TOTAL "||1500|15"}},
        GridCase{"collapse_item_compact.xlsx",
                 {kCompactHeader, "East|500|5", "North|300|3", "South|700|7", TOTAL "|1500|15"}},
        GridCase{"collapse_field_outline.xlsx",
                 {kTabularHeader, "East||500|5", "North||300|3", "South||700|7", TOTAL "||1500|15"}},
        GridCase{"expand_field_outline.xlsx",
                 {kTabularHeader, "East||500|5", "|A|500|5", "North||300|3", "|A|100|1", "|B|200|2", "South||700|7",
                  "|A|300|3", "|B|400|4", TOTAL "||1500|15"}},
        GridCase{"item_caption.xlsx", {kCompactHeader, "zz|500|5", "North|300|3", "South|700|7", TOTAL "|1500|15"}},
        GridCase{"date_field_ungrouped.xlsx",
                 {ROW_LABELS "|\xE5\x90\x88\xE8\xA8\x88 / Amt",  // 合計 / Amt
                  "2024/1/15|100", "2024/1/25|80", "2024/2/8|150", "2024/2/20|50", "2024/3/10|200", TOTAL "|580"},
                 kDateSheet,
                 kDateTable},
        // A blank line on a column field leaves the report unchanged.
        GridCase{"blank_row_outer_column.xlsx",
                 ColumnReport(Line({"East", "", "East Sum of Amt", "East Sum of Qty", "North", "", "", "",
                                    "North Sum of Amt", "North Sum of Qty", "South", "", "", "", "South Sum of Amt",
                                    "South Sum of Qty", ALL_OF "Sum of Amt", ALL_OF "Sum of Qty"}),
                              Line({"A", "", "", "", "A", "", "B", "", "", "", "A", "", "B", "", "", "", "", ""}),
                              Line({"Sum of Amt", "Sum of Qty", "", "", "Sum of Amt", "Sum of Qty", "Sum of Amt",
                                    "Sum of Qty", "", "", "Sum of Amt", "Sum of Qty", "Sum of Amt", "Sum of Qty", "",
                                    "", "", ""}),
                              Line({"500", "5", "500", "5", "100", "1", "200", "2", "300", "3", "300", "3", "400", "4",
                                    "700", "7", "1500", "15"}))},
        GridCase{"repeat_labels_outer_column.xlsx",
                 ColumnReport(Line({"East", "East", "East Sum of Amt", "East Sum of Qty", "North", "North", "North",
                                    "North", "North Sum of Amt", "North Sum of Qty", "South", "South", "South", "South",
                                    "South Sum of Amt", "South Sum of Qty", ALL_OF "Sum of Amt", ALL_OF "Sum of Qty"}),
                              Line({"A", "", "", "", "A", "", "B", "", "", "", "A", "", "B", "", "", "", "", ""}),
                              Line({"Sum of Amt", "Sum of Qty", "", "", "Sum of Amt", "Sum of Qty", "Sum of Amt",
                                    "Sum of Qty", "", "", "Sum of Amt", "Sum of Qty", "Sum of Amt", "Sum of Qty", "",
                                    "", "", ""}),
                              Line({"500", "5", "500", "5", "100", "1", "200", "2", "300", "3", "300", "3", "400", "4",
                                    "700", "7", "1500", "15"}))},
        GridCase{
            "collapse_item_outer_column.xlsx",
            ColumnReport(Line({"East", "", "East Sum of Amt", "East Sum of Qty", "North", "", "South", "", "", "",
                               "South Sum of Amt", "South Sum of Qty", ALL_OF "Sum of Amt", ALL_OF "Sum of Qty"}),
                         Line({"A", "", "", "", "", "", "A", "", "B", "", "", "", "", ""}),
                         Line({"Sum of Amt", "Sum of Qty", "", "", "Sum of Amt", "Sum of Qty", "Sum of Amt",
                               "Sum of Qty", "Sum of Amt", "Sum of Qty", "", "", "", ""}),
                         Line({"500", "5", "500", "5", "300", "3", "300", "3", "400", "4", "700", "7", "1500", "15"}))},
        // Without a grand total row the last block still ends in a blank row.
        GridCase{"blank_row_outer_tabular_no_grand_total.xlsx",
                 {kTabularHeader, "East|A|500|5", "East" SUBTOTAL "||500|5", "|||", "North|A|100|1", "|B|200|2",
                  "North" SUBTOTAL "||300|3", "|||", "South|A|300|3", "|B|400|4", "South" SUBTOTAL "||700|7", "|||"}},
        // Outer and middle blocks ending together each keep their blank row.
        GridCase{"blank_row_outer_middle_tabular.xlsx",
                 {"Region|Product|Qty|Sum of Amt|Sum of Qty",
                  "East|A|5|500|5",
                  "|A" SUBTOTAL "||500|5",
                  "||||",
                  "East" SUBTOTAL "|||500|5",
                  "||||",
                  "North|A|1|100|1",
                  "|A" SUBTOTAL "||100|1",
                  "||||",
                  "|B|2|200|2",
                  "|B" SUBTOTAL "||200|2",
                  "||||",
                  "North" SUBTOTAL "|||300|3",
                  "||||",
                  "South|A|3|300|3",
                  "|A" SUBTOTAL "||300|3",
                  "||||",
                  "|B|4|400|4",
                  "|B" SUBTOTAL "||400|4",
                  "||||",
                  "South" SUBTOTAL "|||700|7",
                  "||||",
                  TOTAL "|||1500|15"}},
        GridCase{"blank_row_outer_compact.xlsx",
                 {ROW_LABELS "|Sum of Amt|Sum of Qty", "East|500|5", "A|500|5", "||", "North|300|3", "A|100|1",
                  "B|200|2", "||", "South|700|7", "A|300|3", "B|400|4", "||", TOTAL "|1500|15"}}));

TEST(PivotOptionsFixtureGrid, DateLabelCellsCarryTheFieldNumberFormat) {
  const std::optional<pivot::PivotCells> cells = ProjectFixture("date_field_ungrouped.xlsx", kDateSheet, kDateTable);
  ASSERT_TRUE(cells.has_value());
  std::vector<std::string> labels;
  std::vector<std::string> formats;
  for (const pivot::PivotCell& cell : cells->cells) {
    if (cell.kind == pivot::PivotCellKind::RowLabel && cell.col == cells->left) {
      labels.push_back(CellText(cell.value));
      formats.push_back(cell.number_format);
    }
  }
  EXPECT_EQ(labels, (std::vector<std::string>{"2024/1/15", "2024/1/25", "2024/2/8", "2024/2/20", "2024/3/10"}));
  EXPECT_EQ(formats, std::vector<std::string>(5, "yyyy/m/d"));
}

}  // namespace
}  // namespace formulon
