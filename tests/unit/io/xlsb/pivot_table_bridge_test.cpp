#include "io/xlsb/pivot_table_bridge.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <functional>
#include <map>
#include <set>
#include <sstream>
#include <string>
#include <utility>
#include <variant>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/pivot_table_reader.h"
#include "io/xlsb/pivot_table_records.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_twin_corpus.h"
#include "sheet.h"
#include "workbook.h"
#include "xlsb_twin_compare.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

namespace tc = test;

constexpr const char* kTablePrefix = "xl/pivotTables/pivotTable";
/// BrtCustomFilter with a string operand: Excel leaves bytes 2-9 uninitialised; the writer emits zeros there.
constexpr std::uint16_t kBrtCustomFilter = 174;
constexpr std::uint8_t kVtsString = 0x06;

/// (xlsb, xlsx) paths of every Excel-saved pivot pair.
std::vector<std::pair<std::string, std::string>> TwinPairs() {
  std::vector<std::pair<std::string, std::string>> pairs;
  for (const std::string& path : pivot_twins::CorpusPaths()) {
    std::string twin = path.substr(0, path.size() - 5U) + ".xlsx";
    if (std::filesystem::exists(twin)) {
      pairs.emplace_back(path, std::move(twin));
    }
  }
  return pairs;
}

/// The host that saved `path`: the fixtures under win/ are Windows Excel saves, the rest Mac Excel saves.
ExcelHost SavingHost(const std::string& path) {
  const std::string root = std::string(FORMULON_FIXTURES_DIR) + "/excel/";
  const bool windows = path == root + "win/pivot_days_end/pivot_days_end.xlsb" ||
                       path == root + "win/pivot_week_1900/pivot_week_1900.xlsb";
  return windows ? ExcelHost::kWin365 : ExcelHost::kMac365;
}

std::string Name(const std::string& path) {
  return std::filesystem::path(path).filename().string();
}

/// Attributes that differ between two saves of one book.
bool IsPerSaveAttribute(const std::string& name) {
  return name == "xr:uid" || name == "cacheId" || name == "updatedVersion" || name == "createdVersion" ||
         name == "minRefreshableVersion";
}

std::string Attrs(std::vector<std::pair<std::string, std::string>> attrs, bool sorted = true) {
  attrs.erase(std::remove_if(attrs.begin(), attrs.end(), [](const auto& a) { return IsPerSaveAttribute(a.first); }),
              attrs.end());
  if (sorted) {
    std::sort(attrs.begin(), attrs.end());
  }
  std::string out;
  for (const auto& a : attrs) {
    out += " " + a.first + "=" + a.second;
  }
  return out;
}

std::string Opt(const std::optional<std::uint32_t>& v) {
  return v ? std::to_string(*v) : "-";
}

std::string FilterValue(const std::variant<std::monostate, int, double, std::string>& v) {
  std::ostringstream out;
  out << v.index() << ":";
  if (const int* i = std::get_if<int>(&v)) {
    out << *i;
  } else if (const double* d = std::get_if<double>(&v)) {
    out.precision(17);
    out << *d;
  } else if (const std::string* t = std::get_if<std::string>(&v)) {
    out << *t;
  }
  return out.str();
}

/// The facts of xlsb_parity_pivot_table_test's PivotLines for one table, then the
/// members the xlsx reader keeps verbatim.
std::vector<std::string> TableLines(const pivot::PivotTable& table) {
  std::vector<std::string> lines;
  const std::string p = "pivot " + table.name() + ": ";
  lines.push_back(p + "anchor " + std::to_string(table.anchor_row()) + "," + std::to_string(table.anchor_col()) +
                  " span " + std::to_string(table.span_rows()) + "x" + std::to_string(table.span_cols()) +
                  " authored " + std::to_string(table.has_authored_span()));
  lines.push_back(p + "location " + Opt(table.location_first_header_row()) + " " +
                  Opt(table.location_first_data_row()) + " " + Opt(table.location_first_data_col()) + " " +
                  Opt(table.location_row_page_count()) + " " + Opt(table.location_col_page_count()));
  lines.push_back(p + "layout " + std::to_string(static_cast<int>(table.layout())) + " data caption " +
                  table.data_caption());
  lines.push_back(p + "grand totals " + std::to_string(table.grand_totals_rows()) +
                  std::to_string(table.grand_totals_cols()) + " caption " + table.grand_total_caption().value_or("-"));
  lines.push_back(
      p + "flags " + std::to_string(table.subtotal_hidden_items()) + std::to_string(table.auto_formatting()) +
      std::to_string(table.enable_drill()) + std::to_string(table.show_error()) + std::to_string(table.show_missing()) +
      std::to_string(table.show_headers()) + std::to_string(table.show_drill()) + std::to_string(table.print_drill()) +
      std::to_string(table.show_data_tips()) + std::to_string(table.grid_drop_zones()) +
      std::to_string(table.show_empty_row()) + std::to_string(table.show_empty_col()) +
      std::to_string(table.field_print_titles()) + std::to_string(table.item_print_titles()) +
      std::to_string(table.show_data_drop_down()) + std::to_string(table.multiple_field_filters()) +
      std::to_string(table.preserve_formatting()) + std::to_string(table.merge_item()) +
      std::to_string(table.page_over_then_down()) + std::to_string(table.custom_list_sort()) + " indent " +
      std::to_string(table.indent()) + " wrap " + std::to_string(table.page_wrap()));
  lines.push_back(p + "captions error " + table.error_caption().value_or("-") + " missing " +
                  table.missing_caption().value_or("-"));
  lines.push_back(p + "alt text " + table.alt_text() + " | " + table.alt_text_summary());
  lines.push_back(p + "attrs" + Attrs(table.passthrough_attrs()));
  lines.push_back(p + "attrs in order" + Attrs(table.passthrough_attrs(), /*sorted=*/false));
  if (table.style_info()) {
    const pivot::PivotStyleInfo& style = *table.style_info();
    lines.push_back(p + "style " + style.name + " " + std::to_string(style.show_row_headers) +
                    std::to_string(style.show_col_headers) + std::to_string(style.show_row_stripes) +
                    std::to_string(style.show_col_stripes) + std::to_string(style.show_last_column));
  } else {
    lines.push_back(p + "style none");
  }
  std::string order = p + "rows";
  for (std::uint32_t f : table.row_field_order()) {
    order += " " + std::to_string(f);
  }
  order += " values@" + (table.row_values_position() ? std::to_string(*table.row_values_position()) : "-");
  order += " cols";
  for (std::uint32_t f : table.col_field_order()) {
    order += " " + std::to_string(f);
  }
  order += " values@" + (table.col_values_position() ? std::to_string(*table.col_values_position()) : "-");
  lines.push_back(order);
  for (const pivot::PivotPageField& page : table.page_fields()) {
    lines.push_back(p + "page " + std::to_string(page.field_index) + " item " + Opt(page.item_index) + " hier " +
                    (page.hierarchy ? std::to_string(*page.hierarchy) : "-") + " " + page.name + " " + page.caption);
  }
  for (std::size_t f = 0; f < table.fields().size(); ++f) {
    const pivot::PivotField& field = table.fields()[f];
    const std::string q = p + "field " + std::to_string(f) + " ";
    lines.push_back(q + "axis " + std::to_string(static_cast<int>(field.axis)) + " name " + field.custom_name +
                    " numfmt " + field.number_format + " subtotal " + std::to_string(field.default_subtotal) +
                    std::to_string(field.subtotal_top) + " show all " + std::to_string(field.show_all) + " blank row " +
                    std::to_string(field.insert_blank_row) + " sort " + std::to_string(field.sort.ascending) +
                    std::to_string(field.sort.manual) + std::to_string(field.sort.stated) + field.sort.by_field);
    std::string fns = q + "subtotal fns";
    for (pivot::SubtotalFn fn : field.subtotal_fns) {
      fns += " " + std::to_string(static_cast<int>(fn));
    }
    lines.push_back(fns);
    lines.push_back(q + "attrs" + Attrs(field.passthrough_attrs, /*sorted=*/false));
    lines.push_back(q + "children " + field.raw_children_xml);
    lines.push_back(q + "x14 fill down " + std::to_string(field.fill_down_labels) + Attrs(field.x14_attrs) + " " +
                    field.x14_children_xml + " ext " + field.raw_ext_xml);
    for (const pivot::PivotItem& item : field.items) {
      lines.push_back(q + "item " + (item.has_cache_index ? std::to_string(item.cache_index) : "-") + " " + item.name +
                      " visible " + std::to_string(item.visible) + " detail " + std::to_string(item.show_detail) +
                      " calc " + std::to_string(item.calculated) + " caption " + item.caption +
                      Attrs(item.passthrough_attrs));
    }
    for (const pivot::PivotMarkerItem& marker : field.marker_items) {
      lines.push_back(q + "marker " + marker.type + " after " + std::to_string(marker.after_items) + " detail " +
                      std::to_string(marker.show_detail) + " " + marker.xml);
    }
  }
  for (const pivot::PivotDataField& data : table.data_fields()) {
    lines.push_back(p + "data " + data.name + " field " + std::to_string(data.field_index) + " agg " +
                    std::to_string(static_cast<int>(data.aggregation)) + " numfmt " + data.number_format + " show as " +
                    std::to_string(static_cast<int>(data.show_as)) + " base " + Opt(data.show_as_base_field) + "/" +
                    Opt(data.show_as_base_item) + Attrs(data.x14_attrs) + " " + data.x14_children_xml + " ext " +
                    data.raw_ext_xml);
  }
  lines.push_back(p + "filters " + std::to_string(table.active_filters().size()) + " retained " +
                  std::to_string(table.retained_filters().size()));
  for (const pivot::PivotFilter& filter : table.active_filters()) {
    lines.push_back(p + "filter " + filter.field_name + " type " + std::to_string(static_cast<int>(filter.type)) +
                    " axis " + std::to_string(static_cast<int>(filter.axis)) + " values " + FilterValue(filter.value) +
                    " " + FilterValue(filter.value_high) + " data field " + std::to_string(filter.data_field_index) +
                    " source " + (filter.source ? filter.source->xml : "-"));
  }
  for (const pivot::RetainedPivotFilter& retained : table.retained_filters()) {
    lines.push_back(p + "retained before " + std::to_string(retained.active_before) + " id " +
                    std::to_string(retained.id) + " " + retained.xml);
  }
  lines.push_back(p + "row items " + table.raw_passthrough_after_row_fields());
  lines.push_back(p + "col items " + table.raw_col_items());
  lines.push_back(p + "before style " + table.raw_passthrough_before_style_info());
  lines.push_back(p + "tail " + table.raw_passthrough_xml());
  return lines;
}

/// Lines present on one side only, capped so one bad table does not flood the log.
std::string LineDiff(const std::vector<std::string>& got, const std::vector<std::string>& want) {
  std::multiset<std::string> a(got.begin(), got.end());
  std::multiset<std::string> b(want.begin(), want.end());
  std::ostringstream out;
  std::size_t shown = 0;
  for (const std::string& line : got) {
    if (b.count(line) == 0U && shown++ < 8U) {
      out << "    xlsb: " << line << "\n";
    }
  }
  for (const std::string& line : want) {
    if (a.count(line) == 0U && shown++ < 16U) {
      out << "    xlsx: " << line << "\n";
    }
  }
  return out.str();
}

/// Each pivot table part of `package` to the worksheet part (`sheetN`) whose rels name it.
std::map<std::string, std::string> PivotSheets(const tc::Bytes& package) {
  const std::string dir = "xl/worksheets/_rels/";
  const std::string target = "../pivotTables/";
  std::map<std::string, std::string> out;
  for (const std::string& name : tc::ListParts(package)) {
    if (name.compare(0, dir.size(), dir) != 0) {
      continue;
    }
    const std::string sheet = name.substr(dir.size(), name.find('.', dir.size()) - dir.size());
    const std::optional<tc::Bytes> rels = tc::ReadPart(package, name);
    const std::string text = rels ? std::string(rels->begin(), rels->end()) : std::string();
    for (std::size_t at = text.find(target); at != std::string::npos; at = text.find(target, at + 1U)) {
      const std::size_t end = text.find('"', at);
      out["xl/" + text.substr(at + 3U, end - at - 3U)] = sheet;
    }
  }
  return out;
}

/// The xlsx twin's pivot tables by name, with the worksheet part each sits on.
struct TwinTable {
  std::size_t sheet_index = 0;
  const pivot::PivotTable* table = nullptr;
  std::string sheet_part;
};

std::map<std::string, TwinTable> TwinTables(const Workbook& wb, const tc::Bytes& xlsx) {
  std::map<std::string, TwinTable> out;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const auto& table : wb.sheet(s).pivot_tables()) {
      EXPECT_EQ(out.count(table->name()), 0U) << "pivot table name repeated: " << table->name();
      out[table->name()] = TwinTable{s, table.get(), std::string()};
    }
  }
  for (const auto& [part, sheet] : PivotSheets(xlsx)) {
    const std::optional<tc::Bytes> bytes = tc::ReadPart(xlsx, part);
    if (!bytes) {
      continue;
    }
    auto table = read_pivot_table_definition(*bytes);
    if (table && out.count(table.value().name()) != 0U) {
      out[table.value().name()].sheet_part = sheet;
    }
  }
  return out;
}

Workbook LoadXlsx(const tc::Bytes& bytes) {
  auto r = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(r.has_value()) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

/// One line per record that differs, the uninitialised string-operand bytes of a custom filter aside.
std::vector<std::string> RecordDiffs(const tc::Bytes& ours, const tc::Bytes& excel) {
  const auto ours_records = tc::SplitPart(ours);
  const auto excel_records = tc::SplitPart(excel);
  std::vector<std::string> diffs;
  if (ours_records.size() != excel_records.size()) {
    diffs.push_back(std::to_string(ours_records.size()) + " records, Excel " + std::to_string(excel_records.size()));
  }
  for (std::size_t i = 0; i < std::min(ours_records.size(), excel_records.size()); ++i) {
    const tc::IndexedRecord& a = ours_records[i];
    const tc::IndexedRecord& b = excel_records[i];
    bool same = a.type == b.type && a.payload.size() == b.payload.size();
    for (std::size_t k = 0; same && k < a.payload.size(); ++k) {
      const bool uninitialised = b.type == kBrtCustomFilter && b.payload[0] == kVtsString && k >= 2U && k < 10U;
      same = a.payload[k] == b.payload[k] || uninitialised;
    }
    if (!same) {
      std::ostringstream line;
      line << "record " << i << " type " << a.type << " (Excel " << b.type << ") ours";
      for (const std::uint8_t byte : a.payload) {
        line << ' ' << std::hex << static_cast<int>(byte);
      }
      line << " | Excel";
      for (const std::uint8_t byte : b.payload) {
        line << ' ' << std::hex << static_cast<int>(byte);
      }
      diffs.push_back(line.str());
      if (diffs.size() >= 8U) {
        break;
      }
    }
  }
  return diffs;
}

std::string Join(const std::vector<std::string>& lines) {
  std::string out;
  for (const std::string& line : lines) {
    out += line + "\n";
  }
  return out;
}

TEST(PivotTableBridge, CorpusIsPresent) {
  EXPECT_EQ(TwinPairs().size(), 84U);
}

// (a) Each Excel-saved table part reads into the model its xlsx twin reads into.
TEST(PivotTableBridge, ReadMatchesTheXlsxTwin) {
  std::size_t parts = 0;
  std::size_t matched = 0;
  std::vector<std::string> failures;
  for (const auto& [xlsb_path, xlsx_path] : TwinPairs()) {
    const tc::Bytes xlsb = tc::ReadBytes(xlsb_path);
    const tc::Bytes xlsx = tc::ReadBytes(xlsx_path);
    const Workbook wb = LoadXlsx(xlsx);
    const std::map<std::string, TwinTable> twins = TwinTables(wb, xlsx);
    const std::map<std::string, std::string> sheets = PivotSheets(xlsb);
    for (const auto& [part, bytes] : pivot_twins::Parts(xlsb_path, kTablePrefix)) {
      ++parts;
      const std::string label = Name(xlsb_path) + " " + part;
      auto read = pivot_table_from_xlsb(SpanOf(bytes));
      if (!read) {
        failures.push_back(label + ": " + read.error().message);
        continue;
      }
      pivot::PivotTable table = std::move(read.value());
      // The one member that names the reader, so it differs from the twin's by design.
      EXPECT_EQ(table.definition_source(), pivot::PivotDefinitionSource::kXlsb) << label;
      const auto twin = twins.find(table.name());
      const auto sheet = sheets.find(part);
      if (twin == twins.end() || sheet == sheets.end() || twin->second.sheet_part != sheet->second) {
        failures.push_back(label + ": no xlsx table " + table.name() + " on the same sheet");
        continue;
      }
      if (const pivot::PivotCache* cache = wb.find_pivot_cache(twin->second.table->pivot_cache_id())) {
        pivot::resolve_pivot_names(table, *cache);
      }
      const std::vector<std::string> got = TableLines(table);
      const std::vector<std::string> want = TableLines(*twin->second.table);
      if (got == want) {
        ++matched;
      } else {
        failures.push_back(label + ":\n" + LineDiff(got, want));
      }
    }
  }
  EXPECT_EQ(parts, 89U);
  EXPECT_TRUE(failures.empty()) << matched << " of " << parts << " parts match\n" << Join(failures);
}

// (b) Each xlsx twin's table writes the records of Excel's own part.
TEST(PivotTableBridge, WriteOfTheXlsxTwinMatchesExcelsPart) {
  std::size_t tables = 0;
  std::size_t matched = 0;
  std::vector<std::string> failures;
  for (const auto& [xlsb_path, xlsx_path] : TwinPairs()) {
    const tc::Bytes xlsx = tc::ReadBytes(xlsx_path);
    const Workbook wb = LoadXlsx(xlsx);
    std::map<std::string, std::vector<std::uint8_t>> excel_parts;
    for (auto& [part, bytes] : pivot_twins::Parts(xlsb_path, kTablePrefix)) {
      auto doc = decode_pivot_table(SpanOf(bytes));
      ASSERT_TRUE(doc.has_value()) << part << ": " << doc.error().message;
      excel_parts[doc.value().view.name] = std::move(bytes);
    }
    for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
      for (const auto& table : wb.sheet(s).pivot_tables()) {
        ++tables;
        const std::string label = Name(xlsb_path) + " " + table->name();
        auto ours = pivot_table_to_xlsb(*table, SavingHost(xlsb_path), s);
        const auto excel = excel_parts.find(table->name());
        if (!ours || excel == excel_parts.end()) {
          failures.push_back(label + ": " + (ours ? "no Excel part" : ours.error().message));
          continue;
        }
        const std::vector<std::string> diffs = RecordDiffs(ours.value(), excel->second);
        if (diffs.empty()) {
          ++matched;
        } else {
          failures.push_back(label + ":\n  " + Join(diffs));
        }
      }
    }
  }
  EXPECT_EQ(tables, 89U);
  EXPECT_TRUE(failures.empty()) << matched << " of " << tables << " tables match\n" << Join(failures);
}

// A table read from an Excel part writes that part back.
TEST(PivotTableBridge, ReadThenWriteReproducesExcelsPart) {
  std::size_t parts = 0;
  std::size_t matched = 0;
  std::vector<std::string> failures;
  for (const std::string& path : pivot_twins::CorpusPaths()) {
    for (const auto& [part, bytes] : pivot_twins::Parts(path, kTablePrefix)) {
      ++parts;
      const std::string label = Name(path) + " " + part;
      auto read = pivot_table_from_xlsb(SpanOf(bytes));
      if (!read) {
        failures.push_back(label + ": " + read.error().message);
        continue;
      }
      auto ours = pivot_table_to_xlsb(read.value(), SavingHost(path), 0);
      if (!ours) {
        failures.push_back(label + ": " + ours.error().message);
        continue;
      }
      const std::vector<std::string> diffs = RecordDiffs(ours.value(), bytes);
      if (diffs.empty()) {
        ++matched;
      } else {
        failures.push_back(label + ":\n  " + Join(diffs));
      }
    }
  }
  EXPECT_TRUE(failures.empty()) << matched << " of " << parts << " parts match\n" << Join(failures);
}

TEST(PivotTableBridge, EditedTableIsNotWritable) {
  const std::string path = pivot_twins::CorpusPaths().front();
  const auto parts = pivot_twins::Parts(path, kTablePrefix);
  ASSERT_FALSE(parts.empty());
  auto read = pivot_table_from_xlsb(SpanOf(parts.front().second));
  ASSERT_TRUE(read.has_value()) << read.error().message;
  pivot::PivotTable table = std::move(read.value());
  table.mark_layout_dirty();
  auto written = pivot_table_to_xlsb(table, ExcelHost::kMac365, 3);
  ASSERT_FALSE(written.has_value());
  EXPECT_EQ(written.error().code, FormulonErrorCode::kIoFeatureNotWritable);
  EXPECT_NE(written.error().context.find("kind=pivot_table"), std::string::npos);
  EXPECT_NE(written.error().context.find("sheet_index=3"), std::string::npos);
}

TEST(PivotTableBridge, FormatsEntryIsNotWritable) {
  const std::string path = pivot_twins::CorpusPaths().front();
  const auto parts = pivot_twins::Parts(path, kTablePrefix);
  ASSERT_FALSE(parts.empty());
  auto read = pivot_table_from_xlsb(SpanOf(parts.front().second));
  ASSERT_TRUE(read.has_value()) << read.error().message;
  pivot::PivotTable table = std::move(read.value());
  ASSERT_TRUE(pivot_table_to_xlsb(table, ExcelHost::kMac365, 0).has_value());
  table.mutable_raw_passthrough_before_style_info() =
      "<formats count=\"1\"><format dxfId=\"0\"><pivotArea type=\"all\"/></format></formats>";
  auto written = pivot_table_to_xlsb(table, ExcelHost::kMac365, 0);
  ASSERT_FALSE(written.has_value());
  EXPECT_EQ(written.error().code, FormulonErrorCode::kIoFeatureNotWritable);
  EXPECT_NE(written.error().context.find("kind=pivot_table"), std::string::npos);
}

TEST(PivotTableBridge, WindowsDescendingFieldIsNotWritable) {
  const std::string path = pivot_twins::CorpusPaths().front();
  const auto parts = pivot_twins::Parts(path, kTablePrefix);
  ASSERT_FALSE(parts.empty());
  auto read = pivot_table_from_xlsb(SpanOf(parts.front().second));
  ASSERT_TRUE(read.has_value()) << read.error().message;
  pivot::PivotTable table = std::move(read.value());
  ASSERT_FALSE(table.fields().empty());
  table.mutable_fields()[0].sort.ascending = false;
  table.mutable_fields()[0].sort.manual = false;
  table.mutable_fields()[0].sort.stated = true;
  EXPECT_TRUE(pivot_table_to_xlsb(table, ExcelHost::kMac365, 0).has_value());
  auto written = pivot_table_to_xlsb(table, ExcelHost::kWin365, 0);
  ASSERT_FALSE(written.has_value());
  EXPECT_EQ(written.error().code, FormulonErrorCode::kIoFeatureNotWritable);
  EXPECT_NE(written.error().message.find("unmeasured"), std::string::npos);
  EXPECT_NE(written.error().context.find("kind=pivot_table"), std::string::npos);
}

/// The first pivot table part of the corpus book whose file name starts with `stem`.
std::vector<std::uint8_t> TablePart(const std::string& stem) {
  for (const std::string& path : pivot_twins::CorpusPaths()) {
    if (Name(path).compare(0, stem.size(), stem) == 0) {
      for (auto& part : pivot_twins::Parts(path, kTablePrefix)) {
        if (part.first == "xl/pivotTables/pivotTable1.bin") {
          return std::move(part.second);
        }
      }
    }
  }
  ADD_FAILURE() << "no pivot table part for " << stem;
  return {};
}

/// Reads `stem`'s part after `edit` changed its decoded records; the read must fail with `code`.
void ExpectReadRefused(const std::string& stem, const std::function<void(PivotTableDoc&)>& edit,
                       FormulonErrorCode code) {
  auto doc = decode_pivot_table(SpanOf(TablePart(stem)));
  ASSERT_TRUE(doc.has_value()) << stem;
  edit(doc.value());
  auto bytes = encode_pivot_table(doc.value());
  ASSERT_TRUE(bytes.has_value()) << stem;
  auto read = pivot_table_from_xlsb(SpanOf(bytes.value()));
  ASSERT_FALSE(read.has_value()) << stem;
  EXPECT_EQ(read.error().code, code) << stem << ": " << read.error().message;
}

/// Writes `stem`'s table after `edit` changed the model; the write must be not writable.
void ExpectWriteRefused(const std::string& stem, const std::function<void(pivot::PivotTable&)>& edit) {
  auto read = pivot_table_from_xlsb(SpanOf(TablePart(stem)));
  ASSERT_TRUE(read.has_value()) << stem << ": " << read.error().message;
  pivot::PivotTable table = std::move(read.value());
  ASSERT_TRUE(pivot_table_to_xlsb(table, ExcelHost::kMac365, 0).has_value()) << stem;
  edit(table);
  auto written = pivot_table_to_xlsb(table, ExcelHost::kMac365, 0);
  ASSERT_FALSE(written.has_value()) << stem;
  EXPECT_EQ(written.error().code, FormulonErrorCode::kIoFeatureNotWritable) << stem;
  EXPECT_NE(written.error().context.find("kind=pivot_table"), std::string::npos) << stem;
}

constexpr FormulonErrorCode kCorrupt = FormulonErrorCode::kIoXlsbRecordCorrupt;
constexpr FormulonErrorCode kUnmeasured = FormulonErrorCode::kIoXlsbUnknownRecord;

void SetU32(std::vector<std::uint8_t>& payload, std::size_t at, std::uint32_t value) {
  for (std::size_t k = 0; k < 4U; ++k) {
    payload[at + k] = static_cast<std::uint8_t>(value >> (8U * k));
  }
}

/// The first BrtBeginSXFilter payload of `doc`.
std::vector<std::uint8_t>& FilterHeader(PivotTableDoc& doc) {
  for (pivot_records::RawRecord& rec : doc.filters) {
    if (rec.type == pivot_records::kBeginSXFilter) {
      return rec.payload;
    }
  }
  ADD_FAILURE() << "no filter";
  return doc.filters.front().payload;
}

TEST(PivotTableBridge, ReadRefusesASecondValuesEntry) {
  ExpectReadRefused(
      "B01_base", [](PivotTableDoc& doc) { doc.col_fields = std::vector<std::uint32_t>{0xFFFFFFFEU, 0xFFFFFFFEU}; },
      kCorrupt);
}

TEST(PivotTableBridge, ReadRefusesAListedFieldOnAnotherAxis) {
  ExpectReadRefused("B01_base", [](PivotTableDoc& doc) { doc.row_fields = std::vector<std::uint32_t>{1U}; }, kCorrupt);
  ExpectReadRefused(
      "B01_base", [](PivotTableDoc& doc) { doc.col_fields = std::vector<std::uint32_t>{0U, 0xFFFFFFFEU}; }, kCorrupt);
  ExpectReadRefused("B17_page_field", [](PivotTableDoc& doc) { (*doc.page_fields)[0].field = 2U; }, kCorrupt);
}

TEST(PivotTableBridge, ReadRefusesADataBitNoDataFieldUses) {
  ExpectReadRefused(
      "B01_base",
      [](PivotTableDoc& doc) { doc.fields[1].axis = static_cast<std::uint8_t>(doc.fields[1].axis | kSxFieldAxisData); },
      kCorrupt);
}

TEST(PivotTableBridge, ReadRefusesADataFieldOverAFieldWithoutTheDataBit) {
  ExpectReadRefused("B01_base", [](PivotTableDoc& doc) { doc.fields[3].axis = 0U; }, kCorrupt);
}

TEST(PivotTableBridge, ReadRefusesLocationOffsetsPastTheTable) {
  ExpectReadRefused(
      "B01_base", [](PivotTableDoc& doc) { doc.location.first_data_row = doc.location.last_row + 1U; }, kCorrupt);
  ExpectReadRefused(
      "B01_base", [](PivotTableDoc& doc) { doc.location.first_data_col = doc.location.last_col + 1U; }, kCorrupt);
}

TEST(PivotTableBridge, ReadRefusesAFilterPastTheFieldsOrDataFields) {
  // BrtBeginSXFilter: fld at bytes 0-3, iMeasureFld (a data-field index) at bytes 20-23.
  ExpectReadRefused("E_ft_v01_eq", [](PivotTableDoc& doc) { SetU32(FilterHeader(doc), 0, 99U); }, kCorrupt);
  ExpectReadRefused("E_ft_v01_eq", [](PivotTableDoc& doc) { SetU32(FilterHeader(doc), 20, 9U); }, kCorrupt);
}

TEST(PivotTableBridge, ReadRefusesTheDefaultSubtotalWithCustomFunctions) {
  ExpectReadRefused(
      "A8_custom_subtotal_max",
      [](PivotTableDoc& doc) {
        for (SxField& field : doc.fields) {
          if ((field.subtotals & ~kSxFieldSubtotalDefault) != 0U) {
            field.subtotals = static_cast<std::uint16_t>(field.subtotals | kSxFieldSubtotalDefault);
          }
        }
      },
      kUnmeasured);
}

TEST(PivotTableBridge, WriteRefusesADataPosition) {
  ExpectWriteRefused("B01_base",
                     [](pivot::PivotTable& t) { t.mutable_passthrough_attrs().emplace_back("dataPosition", "0"); });
}

TEST(PivotTableBridge, WriteRefusesABlockedTable) {
  ExpectWriteRefused("B01_base", [](pivot::PivotTable& t) { t.set_blocked(pivot::BlockedReport{}); });
}

TEST(PivotTableBridge, WriteRefusesAPageFieldNameOrCaption) {
  ExpectWriteRefused("B17_page_field", [](pivot::PivotTable& t) { t.mutable_page_fields()[0].name = "Region"; });
  ExpectWriteRefused("B17_page_field", [](pivot::PivotTable& t) { t.mutable_page_fields()[0].caption = "Region"; });
}

TEST(PivotTableBridge, WriteRefusesAFieldExtension) {
  ExpectWriteRefused("B01_base", [](pivot::PivotTable& t) { t.mutable_fields()[0].fill_down_labels = true; });
  ExpectWriteRefused(
      "B01_base", [](pivot::PivotTable& t) { t.mutable_data_fields()[0].x14_attrs.emplace_back("sourceField", "1"); });
}

TEST(PivotTableBridge, UnmeasuredHeaderBitIsRefusedOnRead) {
  const std::string path = pivot_twins::CorpusPaths().front();
  const auto parts = pivot_twins::Parts(path, kTablePrefix);
  ASSERT_FALSE(parts.empty());
  auto doc = decode_pivot_table(SpanOf(parts.front().second));
  ASSERT_TRUE(doc.has_value());
  doc.value().view.flags4 = 0x00;
  auto bytes = encode_pivot_table(doc.value());
  ASSERT_TRUE(bytes.has_value());
  auto read = pivot_table_from_xlsb(SpanOf(bytes.value()));
  ASSERT_FALSE(read.has_value());
  EXPECT_EQ(read.error().code, FormulonErrorCode::kIoXlsbUnknownRecord);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
