// Round trip of the pivot `<filters>` block, and the load -> save sweep over
// every `<pivotTableDefinition>` attribute and child the reader keeps.

#include <cstddef>
#include <cstdint>
#include <map>
#include <string>
#include <string_view>
#include <utility>
#include <variant>
#include <vector>

#include "gtest/gtest.h"
#include "io/pivot_table_reader.h"
#include "io/pivot_table_writer.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_table_reader_test_helpers.h"
#include "pugixml.hpp"

namespace formulon::io {
namespace {
using namespace pivot_table_reader_test_support;

// Region on rows, an unplaced Date field, and Amount summed twice.
pivot::PivotTable FilterTable() {
  pivot::PivotTable table;
  table.set_name("P");
  table.set_pivot_cache_id(1U);
  table.set_anchor(0U, 0U, 4U, 3U);
  const std::pair<const char*, pivot::PivotAxis> fields[] = {
      {"Region", pivot::PivotAxis::Row}, {"Date", pivot::PivotAxis::None}, {"Amount", pivot::PivotAxis::Value}};
  for (const auto& [name, axis] : fields) {
    pivot::PivotField field;
    field.source_name = name;
    field.axis = axis;
    table.mutable_fields().push_back(std::move(field));
  }
  table.mutable_row_field_order().push_back(0U);
  for (const char* name : {"Sum of Amount", "Max of Amount"}) {
    pivot::PivotDataField data;
    data.name = name;
    data.field_index = 2U;
    table.mutable_data_fields().push_back(std::move(data));
  }
  table.set_col_values_position(0U);
  return table;
}

using Payload = std::variant<std::monostate, int, double, std::string>;

pivot::PivotFilter MakeFilter(pivot::FilterType type, std::string field, Payload value = {}, Payload high = {},
                              std::uint32_t data_field = 0U) {
  pivot::PivotFilter filter;
  filter.axis = pivot::PivotAxis::Row;
  filter.field_name = std::move(field);
  filter.type = type;
  filter.value = std::move(value);
  filter.value_high = std::move(high);
  filter.data_field_index = data_field;
  return filter;
}

std::string FiltersBlock(const std::string& xml) {
  const std::size_t begin = xml.find("<filters");
  const std::size_t end = xml.rfind("</filters>");
  if (begin == std::string::npos || end == std::string::npos) {
    return {};
  }
  return xml.substr(begin, end + std::string_view("</filters>").size() - begin);
}

// `<filter>` on field `fld` with id 1: `attrs` after the id, `criteria` inside
// `<filterColumn>`.
std::string Generated(const char* fld, const char* type, const std::string& attrs, const std::string& criteria) {
  return std::string("<filter fld=\"") + fld + "\" type=\"" + type + "\" evalOrder=\"-1\" id=\"1\"" + attrs +
         "><autoFilter ref=\"A1\"><filterColumn colId=\"0\">" + criteria + "</filterColumn></autoFilter></filter>";
}

std::string Custom(const std::string& filters, bool and_join = false) {
  return std::string(and_join ? "<customFilters and=\"1\">" : "<customFilters>") + filters + "</customFilters>";
}

std::string CustomFilter(const char* op, const char* val) {
  return std::string("<customFilter") + (op[0] == '\0' ? "" : std::string(" operator=\"") + op + "\"") + " val=\"" +
         val + "\"/>";
}

struct GeneratedCase {
  const char* label;
  pivot::PivotFilter filter;
  std::string expected;
};

// Every type is written as Excel writes it from the filter dialogs (the
// spellings of the Mac Excel 365 re-saves, bar the `id`), and reads back into
// `active_filters()` unchanged.
TEST(PivotFiltersRoundTrip, EachTypeWritesExcelsSpellingAndReadsBack) {
  using pivot::FilterType;
  const std::string m0 = " iMeasureFld=\"0\"";
  const GeneratedCase cases[] = {
      {"contains", MakeFilter(FilterType::LabelContains, "Region", std::string("or")),
       Generated("0", "captionContains", " stringValue1=\"or\"", Custom(CustomFilter("", "*or*")))},
      {"contains star", MakeFilter(FilterType::LabelContains, "Region", std::string("*")),
       Generated("0", "captionContains", " stringValue1=\"*\"", Custom(CustomFilter("", "***")))},
      {"begins", MakeFilter(FilterType::LabelBeginsWith, "Region", std::string("a&b")),
       Generated("0", "captionBeginsWith", " stringValue1=\"a&amp;b\"", Custom(CustomFilter("", "a&amp;b*")))},
      {"equal", MakeFilter(FilterType::LabelEqual, "Region", std::string("Apple")),
       Generated("0", "captionEqual", " stringValue1=\"Apple\"", "<filters><filter val=\"Apple\"/></filters>")},
      {"equal escaped", MakeFilter(FilterType::LabelEqual, "Region", std::string("~*star")),
       Generated("0", "captionEqual", " stringValue1=\"~*star\"", "<filters><filter val=\"*star\"/></filters>")},
      {"equal pattern", MakeFilter(FilterType::LabelEqual, "Region", std::string("A*")),
       Generated("0", "captionEqual", " stringValue1=\"A*\"", Custom(CustomFilter("", "A*")))},
      {"equal number", MakeFilter(FilterType::LabelEqual, "Region", std::string("10")),
       Generated("0", "captionEqual", " stringValue1=\"10\"", Custom(CustomFilter("", "10")))},
      {"not equal", MakeFilter(FilterType::LabelNotEqual, "Region", std::string("Apple")),
       Generated("0", "captionNotEqual", " stringValue1=\"Apple\"", Custom(CustomFilter("notEqual", "Apple")))},
      {"not begins", MakeFilter(FilterType::LabelNotBeginsWith, "Region", std::string("A")),
       Generated("0", "captionNotBeginsWith", " stringValue1=\"A\"", Custom(CustomFilter("notEqual", "A*")))},
      {"ends", MakeFilter(FilterType::LabelEndsWith, "Region", std::string("e")),
       Generated("0", "captionEndsWith", " stringValue1=\"e\"", Custom(CustomFilter("", "*e")))},
      {"not ends", MakeFilter(FilterType::LabelNotEndsWith, "Region", std::string("e")),
       Generated("0", "captionNotEndsWith", " stringValue1=\"e\"", Custom(CustomFilter("notEqual", "*e")))},
      {"not contains", MakeFilter(FilterType::LabelNotContains, "Region", std::string("p")),
       Generated("0", "captionNotContains", " stringValue1=\"p\"", Custom(CustomFilter("notEqual", "*p*")))},
      {"caption gt", MakeFilter(FilterType::LabelGreaterThan, "Region", std::string("Banana")),
       Generated("0", "captionGreaterThan", " stringValue1=\"Banana\"", Custom(CustomFilter("greaterThan", "Banana")))},
      {"caption ge", MakeFilter(FilterType::LabelGreaterThanOrEqual, "Region", std::string("Banana")),
       Generated("0", "captionGreaterThanOrEqual", " stringValue1=\"Banana\"",
                 Custom(CustomFilter("greaterThanOrEqual", "Banana")))},
      {"caption lt", MakeFilter(FilterType::LabelLessThan, "Region", std::string("Banana")),
       Generated("0", "captionLessThan", " stringValue1=\"Banana\"", Custom(CustomFilter("lessThan", "Banana")))},
      {"caption le", MakeFilter(FilterType::LabelLessThanOrEqual, "Region", std::string("Banana")),
       Generated("0", "captionLessThanOrEqual", " stringValue1=\"Banana\"",
                 Custom(CustomFilter("lessThanOrEqual", "Banana")))},
      {"caption between", MakeFilter(FilterType::LabelBetween, "Region", std::string("Apricot"), std::string("Cherry")),
       Generated(
           "0", "captionBetween", " stringValue1=\"Apricot\" stringValue2=\"Cherry\"",
           Custom(CustomFilter("greaterThanOrEqual", "Apricot") + CustomFilter("lessThanOrEqual", "Cherry"), true))},
      {"caption not between",
       MakeFilter(FilterType::LabelNotBetween, "Region", std::string("Apricot"), std::string("Cherry")),
       Generated("0", "captionNotBetween", " stringValue1=\"Apricot\" stringValue2=\"Cherry\"",
                 Custom(CustomFilter("lessThan", "Apricot") + CustomFilter("greaterThan", "Cherry")))},
      {"top", MakeFilter(FilterType::ValueTop10, "Region", 2),
       Generated("0", "count", m0, "<top10 val=\"2\" filterVal=\"2\"/>")},
      {"bottom", MakeFilter(FilterType::ValueBottomCount, "Region", 2),
       Generated("0", "count", m0, "<top10 top=\"0\" val=\"2\" filterVal=\"2\"/>")},
      {"top percent", MakeFilter(FilterType::ValueTopPercent, "Region", 50),
       Generated("0", "percent", m0, "<top10 percent=\"1\" val=\"50\" filterVal=\"50\"/>")},
      {"bottom percent", MakeFilter(FilterType::ValueBottomPercent, "Region", 20),
       Generated("0", "percent", m0, "<top10 top=\"0\" percent=\"1\" val=\"20\" filterVal=\"20\"/>")},
      {"top sum", MakeFilter(FilterType::ValueTopSum, "Region", 700),
       Generated("0", "sum", m0, "<top10 val=\"700\" filterVal=\"700\"/>")},
      {"bottom sum", MakeFilter(FilterType::ValueBottomSum, "Region", 350),
       Generated("0", "sum", m0, "<top10 top=\"0\" val=\"350\" filterVal=\"350\"/>")},
      {"value equal", MakeFilter(FilterType::ValueEqual, "Region", 300.0),
       Generated("0", "valueEqual", m0, Custom(CustomFilter("", "300")))},
      {"value not equal", MakeFilter(FilterType::ValueNotEqual, "Region", 300.0),
       Generated("0", "valueNotEqual", m0, Custom(CustomFilter("notEqual", "300")))},
      {"greater", MakeFilter(FilterType::ValueGreaterThan, "Region", 100.0, {}, 1U),
       Generated("0", "valueGreaterThan", " iMeasureFld=\"1\"", Custom(CustomFilter("greaterThan", "100")))},
      {"value ge", MakeFilter(FilterType::ValueGreaterThanOrEqual, "Region", 300.0),
       Generated("0", "valueGreaterThanOrEqual", m0, Custom(CustomFilter("greaterThanOrEqual", "300")))},
      {"value lt", MakeFilter(FilterType::ValueLessThan, "Region", 300.0),
       Generated("0", "valueLessThan", m0, Custom(CustomFilter("lessThan", "300")))},
      {"value le", MakeFilter(FilterType::ValueLessThanOrEqual, "Region", 300.0),
       Generated("0", "valueLessThanOrEqual", m0, Custom(CustomFilter("lessThanOrEqual", "300")))},
      {"between", MakeFilter(FilterType::ValueBetween, "Region", 100.0, 175.5),
       Generated("0", "valueBetween", m0,
                 Custom(CustomFilter("greaterThanOrEqual", "100") + CustomFilter("lessThanOrEqual", "175.5"), true))},
      {"value not between", MakeFilter(FilterType::ValueNotBetween, "Region", 100.0, 300.0),
       Generated("0", "valueNotBetween", m0,
                 Custom(CustomFilter("lessThan", "100") + CustomFilter("greaterThan", "300")))},
      {"date", MakeFilter(FilterType::LabelDate, "Date", 45292.0, 45382.0),
       Generated("1", "dateBetween", "",
                 Custom(CustomFilter("greaterThanOrEqual", "45292") + CustomFilter("lessThanOrEqual", "45382"), true))},
      {"date equal", MakeFilter(FilterType::DateEqual, "Date", 45330.0),
       Generated("1", "dateEqual", "", Custom(CustomFilter("", "45330")))},
      {"date not equal", MakeFilter(FilterType::DateNotEqual, "Date", 45330.0),
       Generated("1", "dateNotEqual", "", Custom(CustomFilter("notEqual", "45330")))},
      {"date older", MakeFilter(FilterType::DateOlderThan, "Date", 45330.0),
       Generated("1", "dateOlderThan", "", Custom(CustomFilter("lessThan", "45330")))},
      {"date older or equal", MakeFilter(FilterType::DateOlderThanOrEqual, "Date", 45330.0),
       Generated("1", "dateOlderThanOrEqual", "", Custom(CustomFilter("lessThanOrEqual", "45330")))},
      {"date newer", MakeFilter(FilterType::DateNewerThan, "Date", 45306.5),
       Generated("1", "dateNewerThan", "", Custom(CustomFilter("greaterThan", "45306.5")))},
      {"date newer or equal", MakeFilter(FilterType::DateNewerThanOrEqual, "Date", 45330.0),
       Generated("1", "dateNewerThanOrEqual", "", Custom(CustomFilter("greaterThanOrEqual", "45330")))},
      {"date not between", MakeFilter(FilterType::DateNotBetween, "Date", 45316.0, 45342.0),
       Generated("1", "dateNotBetween", "",
                 Custom(CustomFilter("lessThan", "45316") + CustomFilter("greaterThan", "45342")))},
      {"this month", MakeFilter(FilterType::DateThisMonth, "Date"),
       Generated("1", "thisMonth", "", "<dynamicFilter type=\"thisMonth\"/>")},
      {"Q2", MakeFilter(FilterType::DateQ2, "Date"), Generated("1", "Q2", "", "<dynamicFilter type=\"Q2\"/>")},
  };
  for (const GeneratedCase& c : cases) {
    SCOPED_TRACE(c.label);
    pivot::PivotTable table = FilterTable();
    table.mutable_active_filters().push_back(c.filter);
    const std::string written = write_pivot_table_definition(table);
    EXPECT_EQ(FiltersBlock(written), "<filters count=\"1\">" + c.expected + "</filters>");
    // Schema order: after the style slot, before the tail extension list.
    EXPECT_LT(written.find("<filters"), written.find("<extLst"));

    auto back_or = read_pivot_table_definition(Bytes(written));
    ASSERT_TRUE(static_cast<bool>(back_or)) << back_or.error().message;
    const pivot::PivotTable& back = back_or.value();
    ASSERT_EQ(back.active_filters().size(), 1U);
    const pivot::PivotFilter& f = back.active_filters()[0];
    EXPECT_EQ(f.type, c.filter.type);
    EXPECT_EQ(f.value, c.filter.value);
    EXPECT_EQ(f.value_high, c.filter.value_high);
    EXPECT_EQ(f.data_field_index, c.filter.data_field_index);
    EXPECT_EQ(f.axis, pivot::PivotAxis::Row);
    ASSERT_TRUE(f.source.has_value());
    EXPECT_EQ(f.source->id, 1U);
    EXPECT_TRUE(back.retained_filters().empty());
    EXPECT_EQ(FiltersBlock(write_pivot_table_definition(back)), FiltersBlock(written));
  }
}

// Every relative period and calendar position reads back as itself.
TEST(PivotFiltersRoundTrip, OperandlessTypesReadBackAsThemselves) {
  for (auto type = static_cast<unsigned>(pivot::FilterType::DateTomorrow);
       type <= static_cast<unsigned>(pivot::FilterType::DateM12); ++type) {
    SCOPED_TRACE(type);
    pivot::PivotTable table = FilterTable();
    table.mutable_active_filters().push_back(MakeFilter(static_cast<pivot::FilterType>(type), "Date"));
    auto back_or = read_pivot_table_definition(Bytes(write_pivot_table_definition(table)));
    ASSERT_TRUE(static_cast<bool>(back_or)) << back_or.error().message;
    ASSERT_EQ(back_or.value().active_filters().size(), 1U);
    EXPECT_EQ(back_or.value().active_filters()[0].type, static_cast<pivot::FilterType>(type));
    EXPECT_TRUE(std::holds_alternative<std::monostate>(back_or.value().active_filters()[0].value));
  }
}

// A generated value filter marks its field `measureFilter="1"`, as Excel does.
TEST(PivotFiltersRoundTrip, AValueFilterMarksItsFieldAsMeasureFiltered) {
  pivot::PivotTable table = FilterTable();
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::ValueGreaterThan, "Region", 150.0));
  const std::string written = write_pivot_table_definition(table);
  EXPECT_NE(written.find("<pivotField axis=\"axisRow\" showAll=\"0\" measureFilter=\"1\""), std::string::npos)
      << written;
  EXPECT_EQ(written.find("measureFilter", written.find("measureFilter") + 1U), std::string::npos) << written;

  pivot::PivotTable caption = FilterTable();
  caption.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::LabelEqual, "Region", std::string("x")));
  EXPECT_EQ(write_pivot_table_definition(caption).find("measureFilter"), std::string::npos);
}

// A value filter naming something off its axis prunes the axis's innermost
// field, so that is the field the file names.
TEST(PivotFiltersRoundTrip, ValueFilterOffItsAxisIsWrittenAgainstTheInnermostField) {
  pivot::PivotTable table = FilterTable();
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::ValueTop10, "Sum of Amount", 3));
  EXPECT_NE(write_pivot_table_definition(table).find("<filter fld=\"0\" type=\"count\""), std::string::npos);
}

// An entry that filters nothing and has no file spelling is not written.
TEST(PivotFiltersRoundTrip, EntriesThatFilterNothingAreNotWritten) {
  pivot::PivotTable table = FilterTable();
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::ValueBetween, "Region", 1.0));
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::LabelContains, "Nowhere", std::string("x")));
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::ValueTop10, "Region", 1, {}, 7U));
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::LabelDate, "Date", std::string("x"), 1.0));
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::LabelBetween, "Region", std::string("a")));
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::LabelEqual, "Region", 1.0));
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::DateEqual, "Date"));
  EXPECT_EQ(write_pivot_table_definition(table).find("<filters"), std::string::npos);
}

std::string MixedFiltersDefinition() {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\" dataCaption=\"Values\">");
  xml.append("<location ref=\"A1:C6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
  xml.append(
      "<pivotFields count=\"3\"><pivotField axis=\"axisRow\" showAll=\"0\"/><pivotField showAll=\"0\"/>"
      "<pivotField dataField=\"1\" showAll=\"0\"/></pivotFields>");
  xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields>");
  xml.append("<dataFields count=\"1\"><dataField name=\"Sum\" fld=\"2\" subtotal=\"sum\"/></dataFields>");
  xml.append(
      "<filters count=\"5\">"
      "<filter fld=\"0\" type=\"captionEqual\" evalOrder=\"-1\" id=\"4\" stringValue1=\"North\"><autoFilter "
      "ref=\"A1\"><filterColumn colId=\"0\"><filters><filter val=\"North\"/></filters></filterColumn></autoFilter>"
      "</filter>"
      "<filter fld=\"0\" type=\"count\" evalOrder=\"-1\" id=\"2\" iMeasureFld=\"0\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><top10 val=\"2\" filterVal=\"2\"/></filterColumn></autoFilter></filter>"
      "<filter fld=\"1\" type=\"someFutureFamily\" evalOrder=\"-1\" id=\"5\" name=\"x\"><autoFilter ref=\"A1\"/>"
      "<extLst><ext uri=\"{00000000-0000-0000-0000-000000000001}\"/></extLst></filter>"
      "<filter fld=\"0\" type=\"percent\" evalOrder=\"-1\" id=\"7\" iMeasureFld=\"0\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><top10 percent=\"1\" val=\"70\" filterVal=\"70\"/></filterColumn></autoFilter>"
      "</filter>"
      "<filter fld=\"1\" type=\"Q2\" evalOrder=\"-1\" id=\"3\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<dynamicFilter type=\"Q2\"/></filterColumn></autoFilter></filter>"
      "</filters>");
  xml.append("</pivotTableDefinition>");
  return xml;
}

// A loaded block keeps every entry, in order and byte for byte: the ones the
// model expresses are listed, the rest are retained in place.
TEST(PivotFiltersRoundTrip, LoadedBlockIsReEmittedVerbatim) {
  const std::string source = MixedFiltersDefinition();
  auto table_or = read_pivot_table_definition(Bytes(source));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.active_filters().size(), 4U);
  EXPECT_EQ(table.active_filters()[0].type, pivot::FilterType::LabelEqual);
  EXPECT_EQ(table.active_filters()[1].type, pivot::FilterType::ValueTop10);
  EXPECT_EQ(table.active_filters()[2].type, pivot::FilterType::ValueTopPercent);
  EXPECT_EQ(table.active_filters()[3].type, pivot::FilterType::DateQ2);
  ASSERT_EQ(table.retained_filters().size(), 1U);
  EXPECT_EQ(table.retained_filters()[0].active_before, 2U);
  EXPECT_EQ(FiltersBlock(write_pivot_table_definition(table)), FiltersBlock(source));
}

// Removing listed entries keeps the retained ones where they were; a new
// entry is appended under an id no loaded entry uses.
TEST(PivotFiltersRoundTrip, EditingTheListKeepsRetainedEntriesInPlace) {
  auto table_or = read_pivot_table_definition(Bytes(MixedFiltersDefinition()));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  pivot::PivotTable& table = table_or.value();
  table.erase_active_filters_if([](const pivot::PivotFilter& f) { return f.type == pivot::FilterType::ValueTop10; });
  EXPECT_EQ(table.retained_filters()[0].active_before, 1U);
  std::string written = write_pivot_table_definition(table);
  std::string block = FiltersBlock(written);
  EXPECT_EQ(block.find("type=\"count\""), std::string::npos) << block;
  EXPECT_EQ(block.rfind("<filters count=\"4\">", 0), 0U) << block;
  EXPECT_LT(block.find("captionEqual"), block.find("someFutureFamily"));
  EXPECT_LT(block.find("someFutureFamily"), block.find("percent"));
  EXPECT_LT(block.find("percent"), block.find("Q2"));

  table.mutable_fields()[0].source_name = "Region";
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::LabelBeginsWith, "Region", std::string("S")));
  written = write_pivot_table_definition(table);
  block = FiltersBlock(written);
  EXPECT_EQ(block.rfind("<filters count=\"5\">", 0), 0U) << block;
  EXPECT_NE(block.find("type=\"captionBeginsWith\" evalOrder=\"-1\" id=\"8\""), std::string::npos) << block;
  EXPECT_LT(block.find("Q2"), block.find("captionBeginsWith"));
}

// A decoded entry whose field moved is re-serialised against the new index
// instead of re-emitting a stale `fld`.
TEST(PivotFiltersRoundTrip, DecodedEntryFollowsItsFieldWhenTheIndexChanges) {
  auto table_or = read_pivot_table_definition(Bytes(MixedFiltersDefinition()));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  pivot::PivotTable& table = table_or.value();
  table.mutable_fields()[0].source_name = "Region";
  table.mutable_fields()[1].source_name = "Date";
  ASSERT_EQ(table.active_filters()[1].type, pivot::FilterType::ValueTop10);
  table.mutable_active_filters()[1].field_name = "Region";
  // Unchanged: the source bytes, `filterVal` included.
  EXPECT_NE(FiltersBlock(write_pivot_table_definition(table)).find("id=\"2\" iMeasureFld=\"0\""), std::string::npos);
  // Region moves to index 1.
  std::swap(table.mutable_fields()[0], table.mutable_fields()[1]);
  table.mutable_row_field_order() = {1U};
  const std::string block = FiltersBlock(write_pivot_table_definition(table));
  EXPECT_NE(block.find("<filter fld=\"1\" type=\"count\" evalOrder=\"-1\" id=\"8\""), std::string::npos) << block;
}

// The shapes Excel writes beyond the dialog's own decode into the model and
// still re-emit their source bytes.
TEST(PivotFiltersRoundTrip, LooselyShapedEntriesDecodeAndReEmitVerbatim) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"2\"><pivotField axis=\"axisRow\"/><pivotField/></pivotFields>");
  xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields>");
  xml.append("<filters count=\"3\">");
  // A pattern that is not `*stringValue1*`: the pattern is what Excel refreshes from.
  xml.append(
      "<filter fld=\"0\" type=\"captionContains\" id=\"1\" stringValue1=\"a\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><customFilters><customFilter val=\"*b*\"/></customFilters></filterColumn>"
      "</autoFilter></filter>");
  // A value filter on a field that sits on no axis.
  xml.append(
      "<filter fld=\"1\" type=\"valueGreaterThan\" id=\"2\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"greaterThan\" val=\"1\"/></customFilters></filterColumn>"
      "</autoFilter></filter>");
  // An equality with only the plain value list: its literal wildcards come back escaped.
  xml.append(
      "<filter fld=\"0\" type=\"captionEqual\" id=\"3\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<filters><filter val=\"*star\"/></filters></filterColumn></autoFilter></filter>");
  xml.append("</filters></pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const std::vector<pivot::PivotFilter>& filters = table_or.value().active_filters();
  ASSERT_EQ(filters.size(), 3U);
  EXPECT_EQ(filters[0].value, Payload(std::string("b")));
  EXPECT_EQ(filters[1].type, pivot::FilterType::ValueGreaterThan);
  EXPECT_EQ(filters[2].value, Payload(std::string("~*star")));
  EXPECT_TRUE(table_or.value().retained_filters().empty());
  EXPECT_EQ(FiltersBlock(write_pivot_table_definition(table_or.value())), FiltersBlock(xml));
}

// Entries the model cannot express stay retained, byte for byte.
TEST(PivotFiltersRoundTrip, UndecodableEntriesStayRetained) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"2\"><pivotField axis=\"axisRow\"/><pivotField/></pivotFields>");
  xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields>");
  xml.append("<filters count=\"7\">");
  // No `fld` at all.
  xml.append("<filter type=\"captionBeginsWith\" id=\"1\" stringValue1=\"a\"><autoFilter ref=\"A1\"/></filter>");
  // A `fld` past the fields.
  xml.append(
      "<filter fld=\"9\" type=\"captionBeginsWith\" id=\"2\" stringValue1=\"a\"><autoFilter ref=\"A1\"/></filter>");
  // A member-property filter.
  xml.append(
      "<filter fld=\"0\" mpFld=\"1\" type=\"captionEqual\" id=\"3\" stringValue1=\"a\"><autoFilter ref=\"A1\"/>"
      "</filter>");
  // An OLAP measure filter.
  xml.append(
      "<filter fld=\"0\" type=\"count\" id=\"4\" iMeasureHier=\"2\"><autoFilter ref=\"A1\"><filterColumn "
      "colId=\"0\"><top10 val=\"2\"/></filterColumn></autoFilter></filter>");
  // A count with no `<top10>`.
  xml.append("<filter fld=\"0\" type=\"count\" id=\"5\"><autoFilter ref=\"A1\"/></filter>");
  // A begins-with pattern without its trailing `*`.
  xml.append(
      "<filter fld=\"0\" type=\"captionBeginsWith\" id=\"6\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter val=\"Hel\"/></customFilters></filterColumn></autoFilter></filter>");
  // A type this version does not know.
  xml.append("<filter fld=\"0\" type=\"someFutureFamily\" id=\"7\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("</filters></pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_TRUE(table_or.value().active_filters().empty());
  EXPECT_EQ(table_or.value().retained_filters().size(), 7U);
  EXPECT_EQ(FiltersBlock(write_pivot_table_definition(table_or.value())), FiltersBlock(xml));
}

// ---------------------------------------------------------------------------
// Attribute sweep
// ---------------------------------------------------------------------------

std::map<std::string, std::string> Attributes(const pugi::xml_node& node) {
  std::map<std::string, std::string> out;
  for (const pugi::xml_attribute attr : node.attributes()) {
    const std::string_view name = attr.name();
    if (name == "xmlns" || name.rfind("xmlns:", 0) == 0) {
      continue;  // Declarations move to wherever the writer binds them.
    }
    out.emplace(attr.name(), attr.value());
  }
  return out;
}

// Same element names, same attribute sets (order-insensitive), same element
// children in the same order.
void ExpectSameTree(const pugi::xml_node& want, const pugi::xml_node& got, const std::string& path) {
  const std::string here = path + "/" + want.name();
  ASSERT_STREQ(want.name(), got.name()) << path;
  EXPECT_EQ(Attributes(want), Attributes(got)) << here;
  std::vector<pugi::xml_node> want_children;
  std::vector<pugi::xml_node> got_children;
  for (const pugi::xml_node child : want.children()) {
    if (child.type() == pugi::node_element) {
      want_children.push_back(child);
    }
  }
  for (const pugi::xml_node child : got.children()) {
    if (child.type() == pugi::node_element) {
      got_children.push_back(child);
    }
  }
  ASSERT_EQ(want_children.size(), got_children.size()) << here;
  for (std::size_t i = 0; i < want_children.size(); ++i) {
    ExpectSameTree(want_children[i], got_children[i], here + "[" + std::to_string(i) + "]");
  }
}

constexpr std::string_view kX14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

// One definition exercising every attribute and child kept across a load ->
// save with no edit in between: modelled state, attribute passthrough on the
// root / fields / items, per-field and table-level report forms that depart
// from the derived one, field children, data-field extensions, the
// position-keyed raw bins, and a mixed `<filters>` block.
std::string SweepDefinition() {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs);
  xml.append(
      " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"xr\" "
      "xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" "
      "xr:uid=\"{11111111-2222-3333-4444-555555555555}\" name=\"PT\" cacheId=\"3\" applyNumberFormats=\"0\" "
      "applyBorderFormats=\"0\" applyFontFormats=\"0\" applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" "
      "applyWidthHeightFormats=\"1\" dataCaption=\"Values\" grandTotalCaption=\"All\" errorCaption=\"#\" "
      "showError=\"1\" missingCaption=\"-\" updatedVersion=\"8\" minRefreshableVersion=\"3\" "
      "useAutoFormatting=\"1\" itemPrintTitles=\"1\" createdVersion=\"8\" indent=\"0\" compact=\"1\" "
      "compactData=\"1\" outline=\"0\" outlineData=\"0\" multipleFieldFilters=\"0\" rowGrandTotals=\"0\" "
      "subtotalHiddenItems=\"1\" chartFormat=\"2\" rowHeaderCaption=\"Rows\">");
  xml.append(
      "<location ref=\"A3:E9\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\" rowPageCount=\"1\" "
      "colPageCount=\"1\"/>");
  xml.append("<pivotFields count=\"5\">");
  xml.append(
      "<pivotField axis=\"axisRow\" compact=\"0\" showAll=\"0\" sortType=\"descending\" defaultSubtotal=\"0\" "
      "insertBlankRow=\"1\" numFmtId=\"14\" name=\"Area\"><items count=\"4\"><item x=\"1\" sd=\"0\"/>"
      "<item t=\"blank\"/><item x=\"0\" n=\"Nord\" h=\"1\"/><item t=\"grand\"/></items><autoSortScope><pivotArea "
      "dataOnly=\"0\" outline=\"0\" "
      "fieldPosition=\"0\"><references count=\"1\"><reference field=\"4294967294\" count=\"1\" selected=\"0\">"
      "<x v=\"0\"/></reference></references></pivotArea></autoSortScope><extLst><ext "
      "uri=\"{2946ED86-A175-432a-8AC1-64E0C546D7DE}\" xmlns:x14=\"");
  xml.append(kX14).append("\"><x14:pivotField fillDownLabels=\"1\"/></ext></extLst></pivotField>");
  xml.append(
      "<pivotField axis=\"axisCol\" compact=\"0\" outline=\"0\" showAll=\"0\" sumSubtotal=\"1\" "
      "countASubtotal=\"1\"><items count=\"4\"><item x=\"0\"/><item x=\"1\"/><item t=\"sum\"/>"
      "<item t=\"countA\"/></items></pivotField>");
  xml.append(
      "<pivotField axis=\"axisPage\" compact=\"0\" showAll=\"0\" sortType=\"ascending\"><items count=\"3\">"
      "<item x=\"0\"/><item x=\"1\" h=\"1\"/><item t=\"default\" sd=\"0\"/></items></pivotField>");
  xml.append("<pivotField dataField=\"1\" compact=\"0\" showAll=\"0\"/>");
  xml.append("<pivotField compact=\"0\" showAll=\"0\" sortType=\"ascending\" dragToRow=\"0\"/>");
  xml.append("</pivotFields>");
  xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields>");
  xml.append("<rowItems count=\"2\"><i><x/></i><i t=\"grand\"><x/></i></rowItems>");
  xml.append("<colFields count=\"2\"><field x=\"1\"/><field x=\"-2\"/></colFields>");
  xml.append("<colItems count=\"1\"><i/></colItems>");
  xml.append("<pageFields count=\"1\"><pageField fld=\"2\" item=\"0\" hier=\"-1\"/></pageFields>");
  xml.append(
      "<dataFields count=\"2\"><dataField name=\"Sum of Amt\" fld=\"3\" subtotal=\"sum\" "
      "showDataAs=\"percentOfTotal\" baseField=\"0\" baseItem=\"0\" numFmtId=\"10\"><extLst>"
      "<ext uri=\"{E15A36E0-9728-4e99-A89B-3F7291B0FE68}\" xmlns:x14=\"");
  xml.append(kX14).append(
      "\"><x14:dataField sourceField=\"0\"/></ext></extLst></dataField>"
      "<dataField name=\"Count of Amt\" fld=\"3\" subtotal=\"count\" baseField=\"0\" baseItem=\"0\"><extLst>"
      "<ext uri=\"{E15A36E0-9728-4e99-A89B-3F7291B0FE68}\" xmlns:x14=\"");
  xml.append(kX14).append(
      "\"><x14:dataField pivotShowAs=\"percentOfParentRow\" sourceField=\"3\" uniqueName=\"[u]\"><x14:extLst>"
      "<x14:ext uri=\"{BBBBBBBB-0000-0000-0000-000000000000}\"/></x14:extLst></x14:dataField></ext><ext "
      "uri=\"{AAAAAAAA-0000-0000-0000-000000000000}\" "
      "xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:dataField "
      "isCountDistinct=\"1\"/></ext></extLst></dataField></dataFields>");
  xml.append(
      "<formats count=\"1\"><format dxfId=\"0\"><pivotArea type=\"all\" dataOnly=\"0\" outline=\"0\" "
      "fieldPosition=\"0\"/></format></formats>");
  xml.append(
      "<conditionalFormats count=\"1\"><conditionalFormat priority=\"1\"><pivotAreas count=\"1\"><pivotArea "
      "type=\"data\" outline=\"0\" fieldPosition=\"0\"/></pivotAreas></conditionalFormat></conditionalFormats>");
  xml.append(
      "<chartFormats count=\"1\"><chartFormat chart=\"0\" format=\"0\" series=\"1\"><pivotArea type=\"data\" "
      "outline=\"0\" fieldPosition=\"0\"/></chartFormat></chartFormats>");
  xml.append(
      "<pivotTableStyleInfo name=\"PivotStyleMedium9\" showRowHeaders=\"1\" showColHeaders=\"1\" "
      "showRowStripes=\"0\" showColStripes=\"1\" showLastColumn=\"1\"/>");
  xml.append(
      "<filters count=\"3\"><filter fld=\"0\" type=\"captionEqual\" evalOrder=\"-1\" id=\"1\" "
      "stringValue1=\"North\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\"><filters><filter val=\"North\"/>"
      "</filters></filterColumn></autoFilter></filter><filter fld=\"0\" type=\"count\" evalOrder=\"-1\" id=\"2\" "
      "iMeasureFld=\"0\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\"><top10 val=\"2\" filterVal=\"2\"/>"
      "</filterColumn></autoFilter></filter><filter fld=\"4\" type=\"Q2\" evalOrder=\"-1\" id=\"3\"><autoFilter "
      "ref=\"A1\"><filterColumn colId=\"0\"><dynamicFilter type=\"Q2\"/></filterColumn></autoFilter></filter>"
      "</filters>");
  xml.append("<rowHierarchiesUsage count=\"1\"><rowHierarchyUsage hierarchyUsage=\"0\"/></rowHierarchiesUsage>");
  xml.append("<extLst><ext uri=\"{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}\" xmlns:x14=\"");
  xml.append(kX14).append(
      "\"><x14:pivotTableDefinition hideValuesRow=\"1\" "
      "xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"/></ext></extLst>");
  xml.append("</pivotTableDefinition>");
  return xml;
}

TEST(PivotDefinitionRoundTrip, LoadSaveKeepsEveryAttributeAndChild) {
  const std::string source = SweepDefinition();
  auto table_or = read_pivot_table_definition(Bytes(source));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const std::string written = write_pivot_table_definition(table_or.value());

  pugi::xml_document want;
  pugi::xml_document got;
  ASSERT_TRUE(want.load_buffer(source.data(), source.size()));
  ASSERT_TRUE(got.load_buffer(written.data(), written.size())) << written;
  ExpectSameTree(want.document_element(), got.document_element(), "");

  auto again_or = read_pivot_table_definition(Bytes(written));
  ASSERT_TRUE(static_cast<bool>(again_or)) << again_or.error().message;
  EXPECT_EQ(write_pivot_table_definition(again_or.value()), written);
}

// A report form read from the file is the file's until the layout is set
// again, which resets the table and every field to the derived form.
TEST(PivotDefinitionRoundTrip, SettingTheLayoutDropsTheFormsReadFromTheFile) {
  auto table_or = read_pivot_table_definition(Bytes(SweepDefinition()));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.layout(), pivot::PivotLayout::Outline);
  table.set_layout(pivot::PivotLayout::Tabular);
  const std::string written = write_pivot_table_definition(table);
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_buffer(written.data(), written.size())) << written;
  const std::map<std::string, std::string> root = Attributes(doc.document_element());
  EXPECT_EQ(root.at("compact"), "0");
  EXPECT_EQ(root.at("compactData"), "0");
  EXPECT_EQ(root.count("outline"), 0U);
  EXPECT_EQ(root.count("outlineData"), 0U);
  std::size_t tabular_fields = 0;
  for (pugi::xml_node field : doc.document_element().child("pivotFields").children("pivotField")) {
    const std::map<std::string, std::string> attrs = Attributes(field);
    EXPECT_EQ(attrs.count("compact") == 1U ? attrs.at("compact") : "", "0") << written;
    EXPECT_EQ(attrs.count("outline") == 1U ? attrs.at("outline") : "", "0") << written;
    ++tabular_fields;
  }
  EXPECT_EQ(tabular_fields, 5U) << written;
}

// A subtotal selection edited in the model regenerates the subtotal markers;
// markers with no model counterpart (`blank`, `grand`) keep their place.
TEST(PivotDefinitionRoundTrip, EditedSubtotalsRegenerateMarkersAndKeepTheRest) {
  auto table_or = read_pivot_table_definition(Bytes(SweepDefinition()));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  pivot::PivotTable& table = table_or.value();
  table.mutable_fields()[0].subtotal_fns = {pivot::SubtotalFn::Max};
  table.mutable_fields()[2].subtotal_fns = {pivot::SubtotalFn::Sum};
  const std::string written = write_pivot_table_definition(table);
  EXPECT_NE(written.find("<items count=\"5\"><item x=\"1\" sd=\"0\"/><item t=\"blank\"/><item x=\"0\" h=\"1\" "
                         "n=\"Nord\"/><item t=\"max\"/><item t=\"grand\"/></items>"),
            std::string::npos)
      << written;
  EXPECT_NE(written.find("<item x=\"1\" h=\"1\"/><item t=\"sum\"/></items>"), std::string::npos) << written;
  // The x14 data-field extension follows a show-as edit and keeps the rest.
  table.mutable_data_fields()[1].show_as = pivot::ShowValuesAs::RankAscending;
  const std::string edited = write_pivot_table_definition(table);
  EXPECT_NE(edited.find("<x14:dataField pivotShowAs=\"rankAscending\" sourceField=\"3\" uniqueName=\"[u]\">"),
            std::string::npos)
      << edited;
}

}  // namespace
}  // namespace formulon::io
