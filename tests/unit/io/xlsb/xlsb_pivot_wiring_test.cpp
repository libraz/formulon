// The XLSB reader and writer on pivots: a cache with its tables is read into
// the model and written back from its retained parts while unchanged, encoded
// from the model and its projected report once edited or refreshed, or left
// alone as a whole.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <map>
#include <memory>
#include <optional>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "cf/cf_types.h"
#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "styles.h"
#include "support/xlsb_package_validator.h"
#include "support/xlsx_package_validator.h"
#include "value.h"
#include "workbook.h"
#include "writer_test_helpers.h"
#include "xlsb_twin_compare.h"
#include "zip_repack_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

namespace tc = test;

std::string Fixture(const std::string& relative) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + relative;
}

Workbook LoadXlsx(const std::string& relative) {
  const tc::Bytes bytes = tc::ReadBytes(Fixture(relative));
  auto r = read_ooxml(SpanOf(bytes));
  EXPECT_TRUE(r.has_value()) << relative << ": " << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto r = read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(r.has_value()) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

std::size_t PivotCount(const Workbook& wb) {
  std::size_t n = 0;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    n += wb.sheet(s).pivot_tables().size();
  }
  return n;
}

/// The `.bin` parts of `package` whose names start with `prefix`, by name.
std::map<std::string, tc::Bytes> Parts(const std::vector<std::uint8_t>& package, const std::string& prefix) {
  std::map<std::string, tc::Bytes> out;
  for (const std::string& name : tc::ListParts(package)) {
    if (name.compare(0, prefix.size(), prefix) == 0 && name.size() > 4U &&
        name.compare(name.size() - 4U, 4U, ".bin") == 0) {
      out[name] = tc::ReadPart(package, name).value_or(tc::Bytes{});
    }
  }
  return out;
}

void ExpectValidXlsb(const std::vector<std::uint8_t>& bytes) {
  const std::vector<std::string> findings = formulon::test::ValidateXlsbPackage(SpanOf(bytes));
  EXPECT_TRUE(findings.empty()) << (findings.empty() ? "" : findings.front());
}

std::string Show(const Value& v) {
  std::ostringstream out;
  out.precision(17);
  if (v.is_text()) {
    out << "s:" << v.as_text();
  } else if (v.is_number()) {
    out << "n:" << v.as_number();
  } else if (v.is_boolean()) {
    out << "b:" << (v.as_boolean() ? 1 : 0);
  } else if (v.is_error()) {
    out << "e:" << static_cast<int>(v.as_error());
  } else {
    out << "blank";
  }
  return out.str();
}

/// Every cache's shared items and records, and every pivot's axes, field items and data fields, as lines.
std::vector<std::string> PivotLines(const Workbook& wb) {
  std::vector<std::string> lines;
  for (const auto& cache : wb.pivot_caches()) {
    lines.push_back("cache records " + std::to_string(cache->records().size()));
    for (const pivot::PivotCacheField& field : cache->fields()) {
      std::string line = "field " + field.name + ":";
      for (const Value& item : field.shared_items) {
        line += " " + Show(item);
      }
      lines.push_back(line);
    }
    for (const pivot::PivotCacheRecord& record : cache->records()) {
      std::string line = "record:";
      for (std::size_t c = 0; c < record.cells.size(); ++c) {
        const Value& v = record.cells[c];
        const bool index = c < record.cell_is_index.size()
                               ? record.cell_is_index[c]
                               : c < cache->fields().size() && !cache->fields()[c].shared_items.empty();
        line += " " + (index && v.is_number()
                           ? Show(cache->fields()[c].shared_items.at(static_cast<std::size_t>(v.as_number())))
                           : Show(v));
      }
      lines.push_back(line);
    }
  }
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const auto& table : wb.sheet(s).pivot_tables()) {
      std::string rows = "pivot " + table->name() + " rows:";
      for (const std::uint32_t f : table->row_field_order()) {
        rows += " " + std::to_string(f);
      }
      lines.push_back(rows);
      for (const auto& field : table->fields()) {
        std::string line = "pivot field:";
        for (const auto& item : field.items) {
          line += " " + std::to_string(item.cache_index) + (item.visible ? "" : "(hidden)");
        }
        lines.push_back(line);
      }
      for (const auto& data : table->data_fields()) {
        lines.push_back("data " + data.name + " " + std::to_string(data.field_index));
      }
    }
  }
  return lines;
}

void SetAltText(Workbook& wb, std::size_t sheet, std::size_t pivot) {
  wb.sheet(sheet).mutable_pivot_tables()[pivot]->set_alt_text("Changed");
}

/// B01_base.xlsx with its one table copied onto the second sheet: two tables, two sheets, one cache.
Workbook TwoTablesOnOneCache() {
  Workbook wb = LoadXlsx("xlsb_pivot_twins/B01_base.xlsx");
  EXPECT_GE(wb.sheet_count(), 2U);
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    if (!wb.sheet(s).pivot_tables().empty()) {
      std::unique_ptr<pivot::PivotTable> copy = wb.sheet(s).pivot_tables().front()->clone();
      copy->set_name("PivotTableCopy");
      wb.sheet(s == 0 ? 1 : 0).add_pivot_table(std::move(copy));
      break;
    }
  }
  EXPECT_EQ(PivotCount(wb), 2U);
  EXPECT_EQ(wb.pivot_caches().size(), 1U);
  return wb;
}

TEST(XlsbPivotWiring, AnXlsxPivotIsWrittenAndReadBack) {
  const Workbook wb = LoadXlsx("xlsb_pivot_twins/B17_page_field.xlsx");
  ASSERT_EQ(PivotCount(wb), 1U);
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 0U);
  for (const auto& [name, bytes] : Parts(saved.value().bytes, "xl/pivot")) {
    EXPECT_FALSE(bytes.empty()) << name;
  }
  EXPECT_EQ(Parts(saved.value().bytes, "xl/pivotTables/").size(), 1U);
  ExpectValidXlsb(saved.value().bytes);
  const Workbook reopened = LoadXlsb(saved.value().bytes);
  EXPECT_EQ(PivotCount(reopened), 1U);
  EXPECT_EQ(reopened.pivot_caches().size(), 1U);
}

// An unchanged group is written from the parts it was read from.
TEST(XlsbPivotWiring, AnUnchangedXlsbPivotIsWrittenFromItsRetainedParts) {
  const tc::Bytes source = tc::ReadBytes(Fixture("xlsb_pivot_twins/B01_base.xlsb"));
  const std::map<std::string, tc::Bytes> source_parts = Parts(source, "xl/pivot");
  ASSERT_EQ(source_parts.size(), 3U);
  const Workbook wb = LoadXlsb(source);
  ASSERT_EQ(PivotCount(wb), 1U);
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 0U);
  for (const auto& [name, bytes] : source_parts) {
    ASSERT_FALSE(bytes.empty()) << name;
    EXPECT_EQ(tc::ReadPart(saved.value().bytes, name), bytes) << name;
  }
  ExpectValidXlsb(saved.value().bytes);
}

// An edited group is encoded from the model instead of failing the save.
TEST(XlsbPivotWiring, AnEditedXlsbPivotIsEncodedFromTheModel) {
  const tc::Bytes source = tc::ReadBytes(Fixture("xlsb_pivot_twins/B01_base.xlsb"));
  Workbook wb = LoadXlsb(source);
  ASSERT_EQ(PivotCount(wb), 1U);
  wb.sheet(0).pivot_tables().empty() ? wb.sheet(1).mutable_pivot_tables()[0]->set_data_caption("Changed")
                                     : wb.sheet(0).mutable_pivot_tables()[0]->set_data_caption("Changed");
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 0U);
  const auto written = Parts(saved.value().bytes, "xl/pivotTables/");
  ASSERT_EQ(written.size(), 1U);
  EXPECT_NE(written.begin()->second, Parts(source, "xl/pivotTables/").begin()->second);
  ExpectValidXlsb(saved.value().bytes);
  const Workbook reopened = LoadXlsb(saved.value().bytes);
  ASSERT_EQ(PivotCount(reopened), 1U);
  for (std::size_t s = 0; s < reopened.sheet_count(); ++s) {
    for (const auto& table : reopened.sheet(s).pivot_tables()) {
      EXPECT_EQ(table->data_caption(), "Changed");
    }
  }
}

// An edited table is encoded with the line items and extent of its projected report, which for an
// unchanged definition are the ones Excel saved.
TEST(XlsbPivotWiring, AnEditedTableIsWrittenFromItsProjection) {
  const tc::Bytes source = tc::ReadBytes(Fixture("xlsb_pivot_twins/B01_base.xlsb"));
  Workbook wb = LoadXlsb(source);
  ASSERT_EQ(PivotCount(wb), 1U);
  std::vector<std::string> read_lines;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (auto& table : wb.sheet(s).mutable_pivot_tables()) {
      read_lines = {table->raw_passthrough_after_row_fields(), table->raw_col_items()};
      table->mark_layout_dirty();
    }
  }
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 0U);
  ExpectValidXlsb(saved.value().bytes);
  const Workbook reopened = LoadXlsb(saved.value().bytes);
  ASSERT_EQ(PivotCount(reopened), 1U);
  EXPECT_EQ(PivotLines(reopened), PivotLines(wb));
  for (std::size_t s = 0; s < reopened.sheet_count(); ++s) {
    for (const auto& table : reopened.sheet(s).pivot_tables()) {
      EXPECT_EQ((std::vector<std::string>{table->raw_passthrough_after_row_fields(), table->raw_col_items()}),
                read_lines);
    }
  }
}

// A cache refreshed from an edited source is written with its tables projected onto the new records.
TEST(XlsbPivotWiring, ARefreshedXlsbPivotIsWrittenFromTheModel) {
  Workbook wb = LoadXlsb(tc::ReadBytes(Fixture("xlsb_pivot_twins/B01_base.xlsb")));
  ASSERT_EQ(PivotCount(wb), 1U);
  // The source is Sheet1!A1:E13 (Region, Product, Date, Amt, Qty) on the second sheet.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(1U, 1U, 0U, "West")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 2U, 3U, Value::number(999.0))));
  for (const auto& cache : wb.pivot_caches()) {
    const auto refreshed = wb.refresh_pivot_cache(cache->cache_id(), eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(refreshed)) << refreshed.error().message;
  }
  const std::vector<std::string> refreshed = PivotLines(wb);
  ASSERT_NE(std::find(refreshed.begin(), refreshed.end(), "field Region: s:West s:North s:East"), refreshed.end());
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 0U);
  ExpectValidXlsb(saved.value().bytes);
  EXPECT_EQ(PivotLines(LoadXlsb(saved.value().bytes)), refreshed);
}

// A cache and table built through the API carry none of the header metadata Excel saves with them (the cache's
// refresh stamps and versions, the table's SXView16 extension and version attributes), so the group is left out.
TEST(XlsbPivotWiring, AnApiBuiltPivotIsLeftOutOfXlsbAndCounted) {
  fm_workbook_t* handle = nullptr;
  ASSERT_EQ(fm_workbook_create_default(&handle), 0);
  const char* const regions[] = {"North", "South", "North", "South"};
  ASSERT_EQ(fm_sheet_set_text(handle, 0U, 0U, 0U, "Region"), 0);
  ASSERT_EQ(fm_sheet_set_text(handle, 0U, 0U, 1U, "Amount"), 0);
  for (std::uint32_t r = 0; r < 4U; ++r) {
    ASSERT_EQ(fm_sheet_set_text(handle, 0U, r + 1U, 0U, regions[r]), 0);
    ASSERT_EQ(fm_sheet_set_number(handle, 0U, r + 1U, 1U, 100.0 * (r + 1U)), 0);
  }
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(handle, 0U, &cache_id), 0) << fm_get_last_error_message();
  fm_pivot_worksheet_source_t source{sizeof(fm_pivot_worksheet_source_t), "A1:B5", "Sheet1", nullptr};
  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(handle, cache_id, &source), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_refresh_pivot_cache(handle, cache_id), 0) << fm_get_last_error_message();
  fm_pivot_spec_t spec{sizeof(fm_pivot_spec_t), "PT", cache_id, 0U, 3U};
  std::size_t pivot = 0;
  ASSERT_EQ(fm_sheet_add_pivot(handle, 0U, &spec, &pivot), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(handle, 0U, pivot, 0U, FM_PIVOT_AXIS_ROW), 0) << fm_get_last_error_message();
  fm_pivot_data_field_t data{};
  data.struct_size = sizeof(fm_pivot_data_field_t);
  data.field_index = 1U;
  data.aggregation = FM_PIVOT_AGGREGATION_SUM;
  std::size_t data_index = 0;
  ASSERT_EQ(fm_sheet_add_pivot_data_field(handle, 0U, pivot, &data, &data_index), 0) << fm_get_last_error_message();

  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  fm_save_diagnostics_info_t info{};
  info.struct_size = sizeof(fm_save_diagnostics_info_t);
  const fm_status_t status = fm_workbook_save_with_diagnostics(handle, FM_WORKBOOK_FORMAT_XLSB, &bytes, &len, &info);
  fm_workbook_destroy(handle);
  ASSERT_EQ(status, 0) << fm_get_last_error_message();
  const std::vector<std::uint8_t> saved(bytes, bytes + len);
  fm_buffer_destroy(bytes);
  EXPECT_EQ(info.deferred_feature_count, 1U);
  EXPECT_TRUE(Parts(saved, "xl/pivot").empty());
  ExpectValidXlsb(saved);
  EXPECT_EQ(PivotCount(LoadXlsb(saved)), 0U);
}

// A refresh keeps the items the source lost in place, so a format area naming an item by position still selects
// it; the .xlsb save writes the area as the .xlsx save does.
TEST(XlsbPivotWiring, ARefreshedTablesFormatAreaKeepsItsItem) {
  Workbook wb = LoadXlsb(tc::ReadBytes(Fixture("xlsb_parity/pivot/formats_areas.xlsb")));
  ASSERT_EQ(PivotCount(wb), 1U);
  // The text of the cache item each format area reference names, through the field's item list.
  const auto selected = [](const Workbook& book) {
    std::vector<std::string> out;
    for (std::size_t s = 0; s < book.sheet_count(); ++s) {
      for (const auto& table : book.sheet(s).pivot_tables()) {
        const pivot::PivotCache* cache = book.find_pivot_cache(table->pivot_cache_id());
        for (const pivot::PivotFormat& format : table->formats()) {
          for (const pivot::PivotTableAreaReference& ref : format.area.references) {
            if (!ref.field.has_value() || *ref.field >= table->fields().size()) {
              continue;
            }
            const std::vector<pivot::PivotItem>& items = table->fields()[*ref.field].items;
            const std::vector<Value>& shared = cache->fields()[*ref.field].shared_items;
            for (const std::uint32_t x : ref.items) {
              out.push_back(x < items.size() && items[x].cache_index < shared.size()
                                ? std::string(shared[items[x].cache_index].as_text())
                                : "out of range");
            }
          }
        }
      }
    }
    return out;
  };
  const std::vector<std::string> before = selected(wb);
  ASSERT_EQ(before, (std::vector<std::string>{"North", "East"}));
  std::size_t source = wb.sheet_count();
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    source = wb.sheet(s).name() == "Sheet1" ? s : source;
  }
  ASSERT_LT(source, wb.sheet_count());
  // Every East row of the source becomes North, so the refreshed cache no longer holds East.
  for (const std::uint32_t row : {1U, 4U, 7U, 10U}) {
    ASSERT_EQ(wb.sheet(source).resolve_cell_value(row, 0U).as_text(), "East");
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(source, row, 0U, "North")));
  }
  for (const auto& cache : wb.pivot_caches()) {
    const auto refreshed = wb.refresh_pivot_cache(cache->cache_id(), eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(refreshed)) << refreshed.error().message;
  }
  EXPECT_EQ(selected(wb), before);

  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 0U);
  ExpectValidXlsb(saved.value().bytes);
  const Workbook from_xlsb = LoadXlsb(saved.value().bytes);
  EXPECT_EQ(selected(from_xlsb), before);
  auto xlsx = write_ooxml(wb);
  ASSERT_TRUE(xlsx.has_value()) << xlsx.error().message;
  auto from_xlsx = read_ooxml(SpanOf(xlsx.value()));
  ASSERT_TRUE(from_xlsx.has_value()) << from_xlsx.error().message;
  EXPECT_EQ(selected(from_xlsx.value().workbook), selected(from_xlsb));
}

// A group the reader cannot model keeps its cache and parts verbatim, as every
// xlsb pivot did before the model could write them.
TEST(XlsbPivotWiring, AnUnreadGroupPassesThroughWhole) {
  const tc::Bytes source = tc::ReadBytes(Fixture("pivot_line_shapes/M6a_page_field_only.xlsb"));
  const std::map<std::string, tc::Bytes> source_parts = Parts(source, "xl/pivot");
  ASSERT_FALSE(source_parts.empty());
  const Workbook wb = LoadXlsb(source);
  EXPECT_EQ(PivotCount(wb), 0U);
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  for (const auto& [name, bytes] : source_parts) {
    EXPECT_EQ(tc::ReadPart(saved.value().bytes, name), bytes) << name;
  }
  ExpectValidXlsb(saved.value().bytes);
}

// An xlsx package cannot carry an unmodelled group's binary parts: the group is left out whole, its
// relationships with it, and counted.
TEST(XlsbPivotWiring, AnXlsxSaveLeavesAnUnreadGroupOutAndCountsIt) {
  const tc::Bytes source = tc::ReadBytes(Fixture("pivot_line_shapes/M6a_page_field_only.xlsb"));
  ASSERT_EQ(Parts(source, "xl/pivotTables/").size(), 1U);
  const Workbook wb = LoadXlsb(source);
  ASSERT_EQ(PivotCount(wb), 0U);
  auto saved = write_ooxml_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 1U);
  for (const std::string& name : tc::ListParts(saved.value().bytes)) {
    EXPECT_EQ(name.find("pivot"), std::string::npos) << name;
  }
  const std::vector<std::string> findings = formulon::test::ValidateXlsxPackage(SpanOf(saved.value().bytes));
  EXPECT_TRUE(findings.empty()) << (findings.empty() ? "" : findings.front());
}

// A table part two sheet relationships name is not modelled; the save keeps it once.
TEST(XlsbPivotWiring, ATablePartRelatedTwiceStaysPassthrough) {
  const tc::Bytes source = tc::ReadBytes(Fixture("xlsb_pivot_twins/B01_base.xlsb"));
  std::string rels_name;
  std::string rels;
  for (const std::string& name : tc::ListParts(source)) {
    const std::optional<tc::Bytes> bytes = tc::ReadPart(source, name);
    const std::string text = bytes ? std::string(bytes->begin(), bytes->end()) : std::string();
    if (name.rfind("xl/worksheets/_rels/", 0) == 0 && text.find("pivotTables/pivotTable1.bin") != std::string::npos) {
      rels_name = name;
      rels = text;
    }
  }
  ASSERT_FALSE(rels_name.empty());
  const std::string extra =
      "<Relationship Id=\"rId99\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
      "pivotTable\" Target=\"../pivotTables/pivotTable1.bin\"/>";
  rels.insert(rels.rfind("</Relationships>"), extra);
  const tc::Bytes doubled = RepackWithEntries(source, {{rels_name, tc::Bytes(rels.begin(), rels.end())}});
  const Workbook wb = LoadXlsb(doubled);
  EXPECT_EQ(PivotCount(wb), 0U);
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(Parts(saved.value().bytes, "xl/pivotTables/").size(), 1U);
  EXPECT_EQ(tc::ReadPart(saved.value().bytes, "xl/pivotTables/pivotTable1.bin"),
            tc::ReadPart(source, "xl/pivotTables/pivotTable1.bin"));
}

TEST(XlsbPivotWiring, TablesOnOneCacheAcrossSheetsAreWrittenAsOneGroup) {
  const Workbook wb = TwoTablesOnOneCache();
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 0U);
  EXPECT_EQ(Parts(saved.value().bytes, "xl/pivotTables/").size(), 2U);
  EXPECT_EQ(Parts(saved.value().bytes, "xl/pivotCache/pivotCacheDefinition").size(), 1U);
  ExpectValidXlsb(saved.value().bytes);
  const Workbook reopened = LoadXlsb(saved.value().bytes);
  EXPECT_EQ(PivotCount(reopened), 2U);
  EXPECT_EQ(reopened.pivot_caches().size(), 1U);
}

// The bridge refusing one table leaves its whole group out: the other table and the cache go with it.
TEST(XlsbPivotWiring, ARefusedTableLeavesItsWholeGroupOut) {
  Workbook wb = TwoTablesOnOneCache();
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    if (!wb.sheet(s).pivot_tables().empty()) {
      // No Excel-saved table carries alternative text, so the table bridge refuses it.
      SetAltText(wb, s, 0);
      break;
    }
  }
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 2U);
  EXPECT_TRUE(Parts(saved.value().bytes, "xl/pivot").empty());
  ExpectValidXlsb(saved.value().bytes);
}

// Three groups in one book, one of them edited: the generated parts are numbered clear of the retained ones.
TEST(XlsbPivotWiring, AnEditedGroupBesideRetainedOnesTakesFreePartNumbers) {
  const tc::Bytes source = tc::ReadBytes(Fixture("win/pivot_days_end/pivot_days_end.xlsb"));
  ASSERT_EQ(Parts(source, "xl/pivotCache/pivotCacheDefinition").size(), 3U);
  Workbook wb = LoadXlsb(source);
  ASSERT_EQ(PivotCount(wb), 3U);
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    if (!wb.sheet(s).pivot_tables().empty()) {
      wb.sheet(s).mutable_pivot_tables()[0]->set_data_caption("Changed");
      break;
    }
  }
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 0U);
  EXPECT_EQ(Parts(saved.value().bytes, "xl/pivotTables/").size(), 3U);
  EXPECT_EQ(Parts(saved.value().bytes, "xl/pivotCache/pivotCacheDefinition").size(), 3U);
  ExpectValidXlsb(saved.value().bytes);
  EXPECT_EQ(PivotCount(LoadXlsb(saved.value().bytes)), 3U);
}

// A cache no table draws from is written from its retained bytes while unchanged.
TEST(XlsbPivotWiring, AnUnchangedCacheWithoutTablesIsKept) {
  const tc::Bytes source = tc::ReadBytes(Fixture("xlsb_pivot_twins/B01_base.xlsb"));
  Workbook wb = LoadXlsb(source);
  ASSERT_EQ(PivotCount(wb), 1U);
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    wb.sheet(s).mutable_pivot_tables().clear();
  }
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 0U);
  const auto definitions = Parts(source, "xl/pivotCache/pivotCacheDefinition");
  ASSERT_EQ(definitions.size(), 1U);
  EXPECT_EQ(tc::ReadPart(saved.value().bytes, definitions.begin()->first), definitions.begin()->second);
  EXPECT_TRUE(Parts(saved.value().bytes, "xl/pivotTables/").empty());
  ExpectValidXlsb(saved.value().bytes);
  const Workbook reopened = LoadXlsb(saved.value().bytes);
  EXPECT_EQ(PivotCount(reopened), 0U);
  EXPECT_EQ(reopened.pivot_caches().size(), 1U);
}

// A cache no table draws from that the bridge refuses is left out and counted.
TEST(XlsbPivotWiring, ARefusedCacheWithoutTablesIsCounted) {
  Workbook wb = LoadXlsx("xlsb_pivot_twins/B01_base.xlsx");
  auto cache = std::make_unique<pivot::PivotCache>();
  cache->set_cache_id(9999U);
  wb.add_pivot_cache(std::move(cache));
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 1U);
  EXPECT_EQ(Parts(saved.value().bytes, "xl/pivotCache/pivotCacheDefinition").size(), 1U);
  ExpectValidXlsb(saved.value().bytes);
}

// A table the bridge writes names model dxfs, so a save that renumbers the dxf table renumbers its formats too.
TEST(XlsbPivotWiring, AnEditedPivotsFormatsFollowTheDxfRenumbering) {
  Workbook wb = LoadXlsb(tc::ReadBytes(Fixture("xlsb_parity/pivot/formats_dxf_order.xlsb")));
  pivot::PivotTable* table = nullptr;
  for (std::size_t s = 0; s < wb.sheet_count() && table == nullptr; ++s) {
    if (!wb.sheet(s).pivot_tables().empty()) {
      table = wb.sheet(s).mutable_pivot_tables()[0].get();
    }
  }
  ASSERT_NE(table, nullptr);
  ASSERT_FALSE(table->formats().empty());
  // Every dxf carries its index as a number-format code, so a renumbered reference can still be resolved.
  std::vector<DifferentialFormat>& dxfs = wb.mutable_styles().dxfs;
  ASSERT_GT(dxfs.size(), 2U);
  for (std::size_t i = 0; i < dxfs.size(); ++i) {
    dxfs[i].has_num_fmt = true;
    dxfs[i].num_fmt_id = static_cast<std::uint16_t>(164U + i);
    dxfs[i].num_fmt_code = std::string(i + 1U, '0');
  }
  std::vector<std::string> want;
  std::uint32_t lowest_format_dxf = 0xFFFFFFFFU;
  for (const pivot::PivotFormat& format : table->formats()) {
    ASSERT_TRUE(format.dxf_id.has_value());
    want.push_back(dxfs[*format.dxf_id].num_fmt_code);
    lowest_format_dxf = std::min(lowest_format_dxf, *format.dxf_id);
  }
  // Dropping a conditional format whose dxf sits below the formats' leaves it unused, so the save moves theirs down.
  bool dropped = false;
  for (std::size_t s = 0; s < wb.sheet_count() && !dropped; ++s) {
    auto& blocks = wb.sheet(s).mutable_conditional_formats();
    for (std::size_t b = 0; b < blocks.size() && !dropped; ++b) {
      for (const auto& rule : blocks[b].rules) {
        if (rule.dxf_id.has_value() && *rule.dxf_id < lowest_format_dxf) {
          blocks.erase(blocks.begin() + static_cast<std::ptrdiff_t>(b));
          dropped = true;
          break;
        }
      }
    }
  }
  ASSERT_TRUE(dropped);
  table->set_data_caption("Changed");
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  const Workbook back = LoadXlsb(saved.value().bytes);
  const pivot::PivotTable* reread = nullptr;
  for (std::size_t s = 0; s < back.sheet_count() && reread == nullptr; ++s) {
    if (!back.sheet(s).pivot_tables().empty()) {
      reread = back.sheet(s).pivot_tables()[0].get();
    }
  }
  ASSERT_NE(reread, nullptr);
  EXPECT_EQ(reread->data_caption(), "Changed");
  EXPECT_LT(back.styles().dxfs.size(), dxfs.size());
  ASSERT_EQ(reread->formats().size(), want.size());
  for (std::size_t f = 0; f < want.size(); ++f) {
    ASSERT_TRUE(reread->formats()[f].dxf_id.has_value());
    EXPECT_NE(*reread->formats()[f].dxf_id, *table->formats()[f].dxf_id) << f;
    ASSERT_LT(*reread->formats()[f].dxf_id, back.styles().dxfs.size());
    EXPECT_EQ(back.styles().dxfs[*reread->formats()[f].dxf_id].num_fmt_code, want[f]) << f;
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
