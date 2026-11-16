// A pivot field sorted by a data field's values, against Excel-saved twins: Region / Product / Amount / Qty on Data,
// one pivot on PV!A3. as1 sorts Region ascending by Sum of Amount, as2 descending, as3 sorts the column field Product
// descending by Sum of Amount, as4 sorts Region descending by Sum of Qty (the second data field), and as5 sorts
// Region by label. The "refreshed" orders are Excel's after Data!C9 went from 35 to -30 and the pivot was refreshed.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "pivot/field_patch.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "workbook.h"
#include "writer_test_helpers.h"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

namespace tc = test;

constexpr std::size_t kData = 0;
constexpr std::size_t kPv = 1;
const char kTablePart[] = "xl/pivotTables/pivotTable1.xml";

tc::Bytes Fixture(const std::string& name) {
  return tc::ReadBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_autosort/" + name);
}

Workbook Load(const std::string& name) {
  const tc::Bytes bytes = Fixture(name);
  if (name.size() > 5U && name.compare(name.size() - 5U, 5U, ".xlsb") == 0) {
    auto read = read_xlsb(SpanOf(bytes));
    EXPECT_TRUE(read.has_value()) << name << ": " << read.error().message;
    return read.has_value() ? std::move(read.value().workbook) : Workbook::create();
  }
  auto read = read_ooxml(SpanOf(bytes));
  EXPECT_TRUE(read.has_value()) << name << ": " << read.error().message;
  return read.has_value() ? std::move(read.value().workbook) : Workbook::create();
}

const pivot::PivotTable& Table(const Workbook& wb) {
  return *wb.sheet(kPv).pivot_tables().at(0);
}

void Refresh(Workbook& wb) {
  ASSERT_FALSE(wb.pivot_caches().empty());
  const auto refreshed = wb.refresh_pivot_cache(wb.pivot_caches()[0]->cache_id(), eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(refreshed)) << refreshed.error().message;
}

/// The labels at `count` cells from (row, col), along a column or a row.
std::vector<std::string> Labels(const Workbook& wb, std::uint32_t row, std::uint32_t col, std::size_t count,
                                bool down) {
  std::vector<std::string> out;
  for (std::size_t i = 0; i < count; ++i) {
    const auto r = static_cast<std::uint32_t>(row + (down ? i : 0U));
    const auto c = static_cast<std::uint32_t>(col + (down ? 0U : i));
    out.emplace_back(wb.sheet(kPv).resolve_cell_value(r, c).as_text());
  }
  return out;
}

std::string Scope(const std::vector<std::uint8_t>& package) {
  const std::optional<tc::Bytes> part = tc::ReadPart(package, kTablePart);
  EXPECT_TRUE(part.has_value());
  const std::string xml = part.has_value() ? std::string(part->begin(), part->end()) : std::string();
  const std::size_t begin = xml.find("<autoSortScope>");
  const std::size_t end = xml.find("</autoSortScope>");
  if (begin == std::string::npos || end == std::string::npos) {
    return std::string();
  }
  EXPECT_EQ(xml.find("<autoSortScope>", begin + 1U), std::string::npos);
  return xml.substr(begin, end + std::string("</autoSortScope>").size() - begin);
}

struct Case {
  const char* name;
  std::size_t field;
  const char* by_field;
  bool ascending;
  // Where the sorted field's labels sit and how many there are.
  std::uint32_t row;
  std::uint32_t col;
  bool down;
  std::vector<std::string> shown;
  std::vector<std::string> refreshed;
};

const std::vector<Case>& Cases() {
  static const std::vector<Case> cases = {
      {"as1",
       0,
       "Sum of Amount",
       true,
       3,
       0,
       true,
       {"North", "West", "East", "South"},
       {"South", "North", "West", "East"}},
      {"as2",
       0,
       "Sum of Amount",
       false,
       3,
       0,
       true,
       {"South", "East", "West", "North"},
       {"East", "West", "North", "South"}},
      {"as3", 1, "Sum of Amount", false, 3, 1, false, {"Fig", "Pear", "Apple"}, {"Pear", "Apple", "Fig"}},
      {"as4",
       0,
       "Sum of Qty",
       false,
       3,
       0,
       true,
       {"North", "East", "West", "South"},
       {"North", "East", "West", "South"}},
      {"as5", 0, "", true, 3, 0, true, {"East", "North", "South", "West"}, {"East", "North", "South", "West"}},
  };
  return cases;
}

TEST(PivotSortScope, ExcelsScopeReadsAsTheSortByDataField) {
  for (const Case& c : Cases()) {
    for (const char* ext : {".xlsx", ".xlsb"}) {
      const Workbook wb = Load(std::string(c.name) + ext);
      const pivot::PivotField& field = Table(wb).fields().at(c.field);
      SCOPED_TRACE(std::string(c.name) + ext);
      EXPECT_EQ(field.sort.by_field, c.by_field);
      EXPECT_FALSE(field.sort.manual);
      EXPECT_EQ(field.sort.ascending, c.ascending);
      EXPECT_TRUE(field.raw_children_xml.empty());
    }
  }
}

TEST(PivotSortScope, TheRenderedOrderFollowsTheValueSortBeforeAndAfterARefresh) {
  for (const Case& c : Cases()) {
    for (const char* ext : {".xlsx", ".xlsb"}) {
      SCOPED_TRACE(std::string(c.name) + ext);
      Workbook wb = Load(std::string(c.name) + ext);
      Refresh(wb);
      EXPECT_EQ(Labels(wb, c.row, c.col, c.shown.size(), c.down), c.shown);
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(kData, 8, 2, Value::number(-30.0))));
      Refresh(wb);
      EXPECT_EQ(Labels(wb, c.row, c.col, c.refreshed.size(), c.down), c.refreshed);
    }
  }
}

TEST(PivotSortScope, AnXlsxSaveWritesExcelsScope) {
  for (const Case& c : Cases()) {
    SCOPED_TRACE(c.name);
    const Workbook wb = Load(std::string(c.name) + ".xlsx");
    auto saved = write_ooxml(wb);
    ASSERT_TRUE(saved.has_value()) << saved.error().message;
    EXPECT_EQ(Scope(saved.value()), Scope(Fixture(std::string(c.name) + ".xlsx")));
  }
}

TEST(PivotSortScope, AnXlsbSaveKeepsTheValueSort) {
  for (const Case& c : Cases()) {
    SCOPED_TRACE(c.name);
    const Workbook wb = Load(std::string(c.name) + ".xlsb");
    auto saved = write_xlsb(wb);
    ASSERT_TRUE(saved.has_value()) << saved.error().message;
    auto back = read_xlsb(SpanOf(saved.value()));
    ASSERT_TRUE(back.has_value()) << back.error().message;
    EXPECT_EQ(Table(back.value().workbook).fields().at(c.field).sort.by_field, c.by_field);
  }
}

// A value sort set on the label-sorted twin saves as Excel saved the same sort.
TEST(PivotSortScope, AValueSortSetThroughThePatchIsSaved) {
  for (const char* ext : {".xlsx", ".xlsb"}) {
    SCOPED_TRACE(ext);
    Workbook wb = Load(std::string("as5") + ext);
    pivot::PivotFieldPatch patch;
    patch.sort_order = pivot::FieldSortOrder::kAscending;
    patch.sort_by_field = "Sum of Amount";
    pivot::PivotTable& table = *wb.sheet(kPv).mutable_pivot_tables().at(0);
    ASSERT_TRUE(static_cast<bool>(pivot::apply_field_patch(wb, table, 0, patch)));
    auto xlsx = write_ooxml(wb);
    ASSERT_TRUE(xlsx.has_value()) << xlsx.error().message;
    EXPECT_EQ(Scope(xlsx.value()), Scope(Fixture("as1.xlsx")));
    auto xlsb = write_xlsb(wb);
    ASSERT_TRUE(xlsb.has_value()) << xlsb.error().message;
    auto back = read_xlsb(SpanOf(xlsb.value()));
    ASSERT_TRUE(back.has_value()) << back.error().message;
    EXPECT_EQ(Table(back.value().workbook).fields().at(0).sort.by_field, "Sum of Amount");
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
