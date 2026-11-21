// Pivot tables in XLSB against Excel's own saves. Every Excel-made `.xlsb`
// pivot that has an `.xlsx` twin (Excel's save of the same book) must load
// into the same pivot model as the twin, keep that model through an `.xlsx`
// save and reopen, and, written back as `.xlsb` from the twin, produce the
// pivot table records Excel wrote.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <map>
#include <set>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/pivot_table_bridge.h"
#include "io/xlsb/pivot_table_records.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "workbook.h"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

namespace tc = test;
using tc::Bytes;

/// (xlsb, xlsx) paths of every Excel-saved pivot pair: the twin corpus, the
/// two Windows-saved grouping books, the two tables with no line items and
/// the two with row fields in different forms.
std::vector<std::pair<std::string, std::string>> TwinPairs() {
  const std::string root = std::string(FORMULON_FIXTURES_DIR) + "/excel/";
  std::vector<std::string> xlsb_paths;
  for (const auto& entry : std::filesystem::directory_iterator(root + "xlsb_pivot_twins")) {
    if (entry.path().extension() == ".xlsb") {
      xlsb_paths.push_back(entry.path().string());
    }
  }
  xlsb_paths.push_back(root + "win/pivot_days_end/pivot_days_end.xlsb");
  xlsb_paths.push_back(root + "win/pivot_week_1900/pivot_week_1900.xlsb");
  xlsb_paths.push_back(root + "pivot_line_shapes/M6a_page_field_only.xlsb");
  xlsb_paths.push_back(root + "pivot_line_shapes/M6b_completely_empty.xlsb");
  xlsb_paths.push_back(root + "pivot_line_shapes/M7a_outer_outline_inner_compact.xlsb");
  xlsb_paths.push_back(root + "pivot_line_shapes/M7b_outer_tabular_inner_outline.xlsb");
  std::sort(xlsb_paths.begin(), xlsb_paths.end());
  std::vector<std::pair<std::string, std::string>> pairs;
  for (const std::string& path : xlsb_paths) {
    std::string twin = path.substr(0, path.size() - 5U) + ".xlsx";
    if (std::filesystem::exists(twin)) {
      pairs.emplace_back(path, std::move(twin));
    }
  }
  return pairs;
}

std::string Name(const std::string& path) {
  return std::filesystem::path(path).filename().string();
}

/// Attributes that differ between two saves of one book.
bool IsPerSaveAttribute(const std::string& name) {
  return name == "xr:uid" || name == "cacheId" || name == "updatedVersion" || name == "createdVersion" ||
         name == "minRefreshableVersion";
}

std::string Attrs(std::vector<std::pair<std::string, std::string>> attrs) {
  attrs.erase(std::remove_if(attrs.begin(), attrs.end(), [](const auto& a) { return IsPerSaveAttribute(a.first); }),
              attrs.end());
  std::sort(attrs.begin(), attrs.end());
  std::string out;
  for (const auto& a : attrs) {
    out += " " + a.first + "=" + a.second;
  }
  return out;
}

std::string Opt(const std::optional<std::uint32_t>& v) {
  return v ? std::to_string(*v) : "-";
}

/// The modelled content of every pivot table, one line per fact, so a
/// mismatch names the fact that differs. The cache a table binds to is named
/// by its position among the workbook's caches, since ids are per save.
std::vector<std::string> PivotLines(const Workbook& wb) {
  std::vector<std::string> lines;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const auto& table : wb.sheet(s).pivot_tables()) {
      const std::string p = "sheet " + std::to_string(s) + " pivot " + table->name() + ": ";
      std::size_t cache_pos = wb.pivot_caches().size();
      for (std::size_t c = 0; c < wb.pivot_caches().size(); ++c) {
        if (wb.pivot_caches()[c]->cache_id() == table->pivot_cache_id()) {
          cache_pos = c;
        }
      }
      lines.push_back(p + "cache #" + std::to_string(cache_pos));
      lines.push_back(p + "anchor " + std::to_string(table->anchor_row()) + "," + std::to_string(table->anchor_col()) +
                      " span " + std::to_string(table->span_rows()) + "x" + std::to_string(table->span_cols()));
      lines.push_back(p + "location " + Opt(table->location_first_header_row()) + " " +
                      Opt(table->location_first_data_row()) + " " + Opt(table->location_first_data_col()) + " " +
                      Opt(table->location_row_page_count()) + " " + Opt(table->location_col_page_count()));
      lines.push_back(p + "layout " + std::to_string(static_cast<int>(table->layout())) + " data caption " +
                      table->data_caption());
      lines.push_back(p + "grand totals " + std::to_string(table->grand_totals_rows()) +
                      std::to_string(table->grand_totals_cols()) + " caption " +
                      table->grand_total_caption().value_or("-"));
      lines.push_back(p + "flags " + std::to_string(table->subtotal_hidden_items()) +
                      std::to_string(table->auto_formatting()) + std::to_string(table->enable_drill()) +
                      std::to_string(table->show_error()) + std::to_string(table->show_missing()) +
                      std::to_string(table->show_headers()) + std::to_string(table->show_drill()) +
                      std::to_string(table->print_drill()) + std::to_string(table->show_data_tips()) +
                      std::to_string(table->grid_drop_zones()) + std::to_string(table->show_empty_row()) +
                      std::to_string(table->show_empty_col()) + std::to_string(table->field_print_titles()) +
                      std::to_string(table->item_print_titles()) + std::to_string(table->show_data_drop_down()) +
                      std::to_string(table->multiple_field_filters()) + std::to_string(table->preserve_formatting()) +
                      std::to_string(table->merge_item()) + std::to_string(table->page_over_then_down()) +
                      std::to_string(table->custom_list_sort()));
      lines.push_back(p + "captions error " + table->error_caption().value_or("-") + " missing " +
                      table->missing_caption().value_or("-"));
      lines.push_back(p + "alt text " + table->alt_text() + " | " + table->alt_text_summary());
      lines.push_back(p + "attrs" + Attrs(table->passthrough_attrs()));
      if (table->style_info()) {
        const pivot::PivotStyleInfo& style = *table->style_info();
        lines.push_back(p + "style " + style.name + " " + std::to_string(style.show_row_headers) +
                        std::to_string(style.show_col_headers) + std::to_string(style.show_row_stripes) +
                        std::to_string(style.show_col_stripes) + std::to_string(style.show_last_column));
      } else {
        lines.push_back(p + "style none");
      }
      std::string order = p + "rows";
      for (std::uint32_t f : table->row_field_order()) {
        order += " " + std::to_string(f);
      }
      order += " values@" + (table->row_values_position() ? std::to_string(*table->row_values_position()) : "-");
      order += " cols";
      for (std::uint32_t f : table->col_field_order()) {
        order += " " + std::to_string(f);
      }
      order += " values@" + (table->col_values_position() ? std::to_string(*table->col_values_position()) : "-");
      lines.push_back(order);
      for (const pivot::PivotPageField& page : table->page_fields()) {
        lines.push_back(p + "page " + std::to_string(page.field_index) + " item " + Opt(page.item_index) + " hier " +
                        (page.hierarchy ? std::to_string(*page.hierarchy) : "-") + " " + page.name);
      }
      for (std::size_t f = 0; f < table->fields().size(); ++f) {
        const pivot::PivotField& field = table->fields()[f];
        const std::string q = p + "field " + std::to_string(f) + " ";
        lines.push_back(q + "axis " + std::to_string(static_cast<int>(field.axis)) + " name " + field.custom_name +
                        " numfmt " + field.number_format + " subtotal " + std::to_string(field.default_subtotal) +
                        std::to_string(field.subtotal_top) + " show all " + std::to_string(field.show_all) + " sort " +
                        std::to_string(field.sort.ascending) + std::to_string(field.sort.manual) +
                        std::to_string(field.sort.stated) + field.sort.by_field);
        std::string fns = q + "subtotal fns";
        for (pivot::SubtotalFn fn : field.subtotal_fns) {
          fns += " " + std::to_string(static_cast<int>(fn));
        }
        lines.push_back(fns);
        lines.push_back(q + "attrs" + Attrs(field.passthrough_attrs));
        for (const pivot::PivotItem& item : field.items) {
          lines.push_back(q + "item " + (item.has_cache_index ? std::to_string(item.cache_index) : "-") + " " +
                          item.name + " visible " + std::to_string(item.visible) + " detail " +
                          std::to_string(item.show_detail) + " calc " + std::to_string(item.calculated) + " caption " +
                          item.caption + Attrs(item.passthrough_attrs));
        }
      }
      for (const pivot::PivotDataField& data : table->data_fields()) {
        lines.push_back(p + "data " + data.name + " field " + std::to_string(data.field_index) + " agg " +
                        std::to_string(static_cast<int>(data.aggregation)) + " numfmt " + data.number_format +
                        " show as " + std::to_string(static_cast<int>(data.show_as)) + " base " +
                        Opt(data.show_as_base_field) + "/" + Opt(data.show_as_base_item) + Attrs(data.x14_attrs));
      }
      lines.push_back(p + "filters " + std::to_string(table->active_filters().size()));
    }
  }
  return lines;
}

/// Lines present on one side only, prefixed "xlsb" / "xlsx" (or the labels
/// given), capped so one bad pair does not flood the log.
std::string LineDiff(const std::vector<std::string>& got, const std::vector<std::string>& want, const char* got_label,
                     const char* want_label) {
  std::multiset<std::string> a(got.begin(), got.end());
  std::multiset<std::string> b(want.begin(), want.end());
  std::ostringstream out;
  std::size_t shown = 0;
  for (const std::string& line : got) {
    if (b.count(line) == 0U && shown++ < 12U) {
      out << "  " << got_label << ": " << line << "\n";
    }
  }
  for (const std::string& line : want) {
    if (a.count(line) == 0U && shown++ < 24U) {
      out << "  " << want_label << ": " << line << "\n";
    }
  }
  return out.str();
}

bool LoadXlsb(const Bytes& bytes, Workbook* out, std::string* error) {
  auto r = read_xlsb(ByteSpan{bytes.data(), bytes.size()});
  if (!r) {
    *error = r.error().message;
    return false;
  }
  *out = std::move(r.value().workbook);
  return true;
}

bool LoadXlsx(const Bytes& bytes, Workbook* out, std::string* error) {
  auto r = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  if (!r) {
    *error = r.error().message;
    return false;
  }
  *out = std::move(r.value().workbook);
  return true;
}

/// BrtCustomFilter with a string operand: Excel leaves bytes 4-9 uninitialised and the writer emits zeros there.
bool SamePayload(std::uint16_t type, const Bytes& excel, const Bytes& ours) {
  constexpr std::uint16_t kBrtCustomFilter = 174;
  constexpr std::uint8_t kVtsString = 0x06;
  if (excel.size() != ours.size()) {
    return false;
  }
  for (std::size_t k = 0; k < excel.size(); ++k) {
    const bool uninitialised = type == kBrtCustomFilter && excel[0] == kVtsString && k >= 4U && k < 10U;
    if (excel[k] != ours[k] && !uninitialised) {
      return false;
    }
  }
  return true;
}

TEST(XlsbParityPivotTableTwins, CorpusIsPresent) {
  // 85 twin pairs, the two Windows grouping books, the two tables with no line items and the two with mixed row
  // field forms.
  EXPECT_EQ(TwinPairs().size(), 93U);
}

TEST(XlsbParityPivotTableTwins, ReadMatchesXlsxTwin) {
  for (const auto& [xlsb_path, xlsx_path] : TwinPairs()) {
    SCOPED_TRACE(Name(xlsb_path));
    Workbook from_xlsb = Workbook::create_empty();
    Workbook from_xlsx = Workbook::create_empty();
    std::string error;
    ASSERT_TRUE(LoadXlsb(tc::ReadBytes(xlsb_path), &from_xlsb, &error)) << error;
    ASSERT_TRUE(LoadXlsx(tc::ReadBytes(xlsx_path), &from_xlsx, &error)) << error;
    const std::vector<std::string> want = PivotLines(from_xlsx);
    ASSERT_FALSE(want.empty());
    const std::vector<std::string> got = PivotLines(from_xlsb);
    EXPECT_EQ(got, want) << LineDiff(got, want, "xlsb", "xlsx");
  }
}

TEST(XlsbParityPivotTableTwins, XlsxSaveOfAnXlsbPivotReopensAsTheTwin) {
  for (const auto& [xlsb_path, xlsx_path] : TwinPairs()) {
    SCOPED_TRACE(Name(xlsb_path));
    Workbook from_xlsb = Workbook::create_empty();
    Workbook from_xlsx = Workbook::create_empty();
    std::string error;
    ASSERT_TRUE(LoadXlsb(tc::ReadBytes(xlsb_path), &from_xlsb, &error)) << error;
    ASSERT_TRUE(LoadXlsx(tc::ReadBytes(xlsx_path), &from_xlsx, &error)) << error;
    auto saved = write_ooxml(from_xlsb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    Workbook reopened = Workbook::create_empty();
    ASSERT_TRUE(LoadXlsx(saved.value(), &reopened, &error)) << error;
    const std::vector<std::string> want = PivotLines(from_xlsx);
    ASSERT_FALSE(want.empty());
    const std::vector<std::string> got = PivotLines(reopened);
    EXPECT_EQ(got, want) << LineDiff(got, want, "reopened", "xlsx");
  }
}

TEST(XlsbParityPivotTableTwins, XlsbSaveOfTheTwinWritesExcelsTableRecords) {
  for (const auto& [xlsb_path, xlsx_path] : TwinPairs()) {
    SCOPED_TRACE(Name(xlsb_path));
    Workbook from_xlsx = Workbook::create_empty();
    std::string error;
    ASSERT_TRUE(LoadXlsx(tc::ReadBytes(xlsx_path), &from_xlsx, &error)) << error;
    // Each twin is encoded as its saving host writes it: the books under win/ by Windows Excel, the rest by Mac.
    const ExcelHost host = xlsb_path.find("/win/") != std::string::npos ? ExcelHost::kWin365 : ExcelHost::kMac365;
    std::map<std::string, Bytes> ours;
    for (std::size_t s = 0; s < from_xlsx.sheet_count(); ++s) {
      for (const auto& table : from_xlsx.sheet(s).pivot_tables()) {
        auto bytes = pivot_table_to_xlsb(*table, host, s);
        ASSERT_TRUE(static_cast<bool>(bytes)) << table->name() << ": " << bytes.error().message;
        ours[table->name()] = std::move(bytes.value());
      }
    }
    const Bytes excel = tc::ReadBytes(xlsb_path);
    std::vector<std::string> diffs;
    for (const std::string& part : tc::ListParts(excel)) {
      if (part.rfind("xl/pivotTables/", 0) != 0 || !tc::IsRecordPart(part)) {
        continue;
      }
      const auto excel_part = tc::ReadPart(excel, part);
      auto doc = decode_pivot_table(ByteSpan{excel_part->data(), excel_part->size()});
      ASSERT_TRUE(static_cast<bool>(doc)) << part << ": " << doc.error().message;
      const auto our_part = ours.find(doc.value().view.name);
      if (our_part == ours.end()) {
        diffs.push_back(part + ": no twin table " + doc.value().view.name);
        continue;
      }
      const auto excel_records = tc::SplitPart(*excel_part);
      const auto our_records = tc::SplitPart(our_part->second);
      if (excel_records.size() != our_records.size()) {
        diffs.push_back(part + ": " + std::to_string(our_records.size()) + " records, Excel " +
                        std::to_string(excel_records.size()));
        continue;
      }
      for (std::size_t i = 0; i < excel_records.size(); ++i) {
        if (excel_records[i].type != our_records[i].type ||
            !SamePayload(excel_records[i].type, excel_records[i].payload, our_records[i].payload)) {
          diffs.push_back(part + ": record " + std::to_string(i) + " type " + std::to_string(our_records[i].type) +
                          ", Excel " + std::to_string(excel_records[i].type));
          break;
        }
      }
    }
    EXPECT_TRUE(diffs.empty()) << [&] {
      std::string all;
      for (const std::string& d : diffs) {
        all += "  " + d + "\n";
      }
      return all;
    }();
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
