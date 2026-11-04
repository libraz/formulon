// Pivot-table definition reader tests grouped by XML concern.

#include <cstddef>
#include <cstdint>
#include <iterator>
#include <string>
#include <string_view>
#include <variant>
#include <vector>

#include "gtest/gtest.h"
#include "io/pivot_table_reader.h"
#include "io/pivot_table_writer.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_table_reader_test_helpers.h"
#include "utils/error.h"

namespace formulon::io {
namespace {
using namespace pivot_table_reader_test_support;

using Payload = std::variant<std::monostate, int, double, std::string>;

// A definition with six fields, field 0 on rows, and `filters` as its
// `<filters>` children.
std::string DefinitionWithFilters(std::string_view filters, std::size_t count) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append(
      "<pivotFields count=\"6\"><pivotField axis=\"axisRow\"/><pivotField/><pivotField/><pivotField/>"
      "<pivotField/><pivotField/></pivotFields>");
  xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields>");
  xml.append("<pivotTableStyleInfo name=\"PivotStyleLight16\"/>");
  xml.append("<filters count=\"").append(std::to_string(count)).append("\">");
  xml.append(filters);
  xml.append("</filters></pivotTableDefinition>");
  return xml;
}

pivot::PivotTable Read(const std::string& xml) {
  auto table_or = read_pivot_table_definition(Bytes(xml));
  EXPECT_TRUE(static_cast<bool>(table_or)) << (table_or ? "" : table_or.error().message);
  return table_or ? std::move(table_or.value()) : pivot::PivotTable{};
}

TEST(PivotTableReader, FiltersDecodeIntoTheListAndRoundTripVerbatim) {
  const std::string xml = DefinitionWithFilters(
      "<filter fld=\"0\" type=\"captionEqual\" evalOrder=\"-1\" id=\"1\">"
      "<autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"equal\" val=\"North\"/></customFilters>"
      "</filterColumn></autoFilter></filter>",
      1U);
  const pivot::PivotTable table = Read(xml);
  ASSERT_EQ(table.active_filters().size(), 1U);
  const pivot::PivotFilter& filter = table.active_filters()[0];
  EXPECT_EQ(filter.type, pivot::FilterType::LabelEqual);
  EXPECT_EQ(filter.value, Payload(std::string("North")));
  ASSERT_TRUE(filter.source.has_value());
  EXPECT_EQ(filter.source->field_index, 0U);
  EXPECT_EQ(filter.source->id, 1U);
  EXPECT_TRUE(filter.field_name.empty());
  EXPECT_TRUE(table.retained_filters().empty());
  EXPECT_EQ(table.raw_passthrough_xml().find("<filters"), std::string::npos);

  const std::string round = write_pivot_table_definition(table);
  EXPECT_EQ(CountOccurrences(round, "<filters"), 1U) << round;
  EXPECT_NE(round.find("val=\"North\""), std::string::npos) << round;
  // The style block is authored before `<filters>`; re-emission keeps that
  // order, which the schema requires.
  EXPECT_LT(round.find("<pivotTableStyleInfo"), round.find("<filters"));
  EXPECT_EQ(write_pivot_table_definition(Read(round)), round);
}

TEST(PivotTableReader, CaptionTypesDecodeTheirOperands) {
  const pivot::PivotTable table = Read(DefinitionWithFilters(
      "<filter fld=\"1\" type=\"captionBeginsWith\"><autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"equal\" val=\"No*\"/></customFilters>"
      "</filterColumn></autoFilter></filter>"
      "<filter fld=\"2\" type=\"captionContains\"><autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"equal\" val=\"*or*\"/></customFilters>"
      "</filterColumn></autoFilter></filter>"
      "<filter fld=\"3\" type=\"captionEndsWith\"><autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"equal\" val=\"*th\"/></customFilters>"
      "</filterColumn></autoFilter></filter>"
      "<filter fld=\"4\" type=\"captionBetween\"><autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters and=\"1\"><customFilter operator=\"greaterThanOrEqual\" val=\"B\"/>"
      "<customFilter operator=\"lessThanOrEqual\" val=\"M\"/></customFilters>"
      "</filterColumn></autoFilter></filter>"
      "<filter fld=\"5\" type=\"captionNotBetween\" stringValue1=\"Apricot\" stringValue2=\"Cherry\">"
      "<autoFilter ref=\"A1\"><filterColumn colId=\"0\"><customFilters><customFilter operator=\"lessThan\" "
      "val=\"Apricot\"/><customFilter operator=\"greaterThan\" val=\"Cherry\"/></customFilters></filterColumn>"
      "</autoFilter></filter>",
      5U));
  const auto& decoded = table.active_filters();
  ASSERT_EQ(decoded.size(), 5U);
  EXPECT_EQ(decoded[0].source->field_index, 1U);
  EXPECT_EQ(decoded[0].type, pivot::FilterType::LabelBeginsWith);
  EXPECT_EQ(decoded[0].value, Payload(std::string("No")));
  EXPECT_EQ(decoded[1].type, pivot::FilterType::LabelContains);
  EXPECT_EQ(decoded[1].value, Payload(std::string("or")));
  EXPECT_EQ(decoded[2].type, pivot::FilterType::LabelEndsWith);
  EXPECT_EQ(decoded[2].value, Payload(std::string("th")));
  EXPECT_EQ(decoded[3].type, pivot::FilterType::LabelBetween);
  EXPECT_EQ(decoded[3].value, Payload(std::string("B")));
  EXPECT_EQ(decoded[3].value_high, Payload(std::string("M")));
  EXPECT_EQ(decoded[4].type, pivot::FilterType::LabelNotBetween);
  EXPECT_EQ(decoded[4].value, Payload(std::string("Apricot")));
  EXPECT_EQ(decoded[4].value_high, Payload(std::string("Cherry")));
}

TEST(PivotTableReader, CaptionEqualAcceptsThePlainFilterListShape) {
  const pivot::PivotTable table =
      Read(DefinitionWithFilters("<filter fld=\"0\" type=\"captionEqual\"><autoFilter ref=\"A3:A9\">"
                                 "<filterColumn colId=\"0\"><filters><filter val=\"South\"/></filters>"
                                 "</filterColumn></autoFilter></filter>",
                                 1U));
  ASSERT_EQ(table.active_filters().size(), 1U);
  EXPECT_EQ(table.active_filters()[0].value, Payload(std::string("South")));
}

// Excel refreshes a begins / ends / contains filter from the nested pattern,
// not from `stringValue1` (beginsWith, stringValue1 `xyz`, pattern `Hel*`
// shows Hello). Without a nested criterion, `stringValue1` is the operand.
// The writer still re-emits `stringValue1` as read.
TEST(PivotTableReader, CaptionWildcardFilterPrefersTheNestedPattern) {
  const pivot::PivotTable table = Read(DefinitionWithFilters(
      "<filter fld=\"0\" type=\"captionBeginsWith\" id=\"1\" stringValue1=\"xyz\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><customFilters><customFilter val=\"Hel*\"/></customFilters>"
      "</filterColumn></autoFilter></filter>"
      "<filter fld=\"0\" type=\"captionContains\" id=\"2\" stringValue1=\"*\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><customFilters><customFilter val=\"*~**\"/></customFilters>"
      "</filterColumn></autoFilter></filter>"
      "<filter fld=\"0\" type=\"captionEndsWith\" id=\"3\" stringValue1=\"?b\"><autoFilter ref=\"A1\"/></filter>"
      "<filter fld=\"0\" type=\"captionEqual\" id=\"4\" stringValue1=\"abc\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><filters><filter val=\"zzz\"/></filters></filterColumn></autoFilter></filter>",
      4U));
  const auto& decoded = table.active_filters();
  ASSERT_EQ(decoded.size(), 4U);
  EXPECT_EQ(decoded[0].value, Payload(std::string("Hel")));
  EXPECT_EQ(decoded[1].value, Payload(std::string("~*")));
  EXPECT_EQ(decoded[2].value, Payload(std::string("?b")));
  // Equality keeps preferring the direct attribute.
  EXPECT_EQ(decoded[3].value, Payload(std::string("abc")));

  const std::string round = write_pivot_table_definition(table);
  EXPECT_NE(round.find("stringValue1=\"xyz\""), std::string::npos) << round;
  EXPECT_NE(round.find("val=\"Hel*\""), std::string::npos) << round;
  EXPECT_NE(round.find("stringValue1=\"?b\""), std::string::npos) << round;
}

TEST(PivotTableReader, UndecodableFiltersAreRetainedAndRoundTrip) {
  const pivot::PivotTable table = Read(DefinitionWithFilters(
      // A count filter with no nested `<top10>` names no quantity at all.
      "<filter fld=\"0\" type=\"count\"><autoFilter ref=\"A1\"/></filter>"
      // An unknown type is the forward-compatibility case: a family this
      // version does not know must be passed through rather than guessed at.
      "<filter fld=\"1\" type=\"someFutureFamily\"><autoFilter ref=\"A1\"/></filter>"
      // A caption filter whose criterion element is missing entirely.
      "<filter fld=\"2\" type=\"captionEqual\"><autoFilter ref=\"A1\"/></filter>",
      3U));
  EXPECT_TRUE(table.active_filters().empty());
  EXPECT_EQ(table.retained_filters().size(), 3U);

  const std::string round = write_pivot_table_definition(table);
  EXPECT_EQ(CountOccurrences(round, "<filter "), 3U) << round;
  EXPECT_NE(round.find("someFutureFamily"), std::string::npos) << round;
}

TEST(PivotTableReader, TopNFlavoursAreDistinguishedByTheirTypeAndDirection) {
  // The flavours share an identically shaped `<top10 val="N">`, so a decode
  // that ignored the type attribute or `top` would silently conflate them.
  const pivot::PivotTable table =
      Read(DefinitionWithFilters("<filter fld=\"0\" type=\"count\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
                                 "<top10 val=\"2\"/></filterColumn></autoFilter></filter>"
                                 "<filter fld=\"0\" type=\"count\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
                                 "<top10 top=\"0\" val=\"2\"/></filterColumn></autoFilter></filter>"
                                 "<filter fld=\"0\" type=\"percent\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
                                 "<top10 percent=\"1\" val=\"70\"/></filterColumn></autoFilter></filter>"
                                 "<filter fld=\"0\" type=\"percent\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
                                 "<top10 top=\"0\" percent=\"1\" val=\"20\"/></filterColumn></autoFilter></filter>"
                                 "<filter fld=\"0\" type=\"sum\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
                                 "<top10 val=\"100\"/></filterColumn></autoFilter></filter>"
                                 "<filter fld=\"0\" type=\"sum\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
                                 "<top10 top=\"0\" val=\"2.5\"/></filterColumn></autoFilter></filter>",
                                 6U));
  const auto& filters = table.active_filters();
  ASSERT_EQ(filters.size(), 6U);
  EXPECT_EQ(filters[0].type, pivot::FilterType::ValueTop10);
  EXPECT_EQ(filters[0].value, Payload(2));
  EXPECT_EQ(filters[1].type, pivot::FilterType::ValueBottomCount);
  EXPECT_EQ(filters[2].type, pivot::FilterType::ValueTopPercent);
  EXPECT_EQ(filters[2].value, Payload(70));
  EXPECT_EQ(filters[3].type, pivot::FilterType::ValueBottomPercent);
  EXPECT_EQ(filters[4].type, pivot::FilterType::ValueTopSum);
  EXPECT_EQ(filters[4].value, Payload(100));
  EXPECT_EQ(filters[5].type, pivot::FilterType::ValueBottomSum);
  EXPECT_EQ(filters[5].value, Payload(2.5));
}

TEST(PivotTableReader, RelativeAndCalendarFiltersDecodeFromTheTypeNameAlone) {
  // These carry no criteria: the period is implied by the type.
  const char* const names[] = {"thisMonth", "yearToDate", "lastQuarter", "thisWeek", "M1", "M12", "Q1", "Q4"};
  const pivot::FilterType types[] = {pivot::FilterType::DateThisMonth,   pivot::FilterType::DateYearToDate,
                                     pivot::FilterType::DateLastQuarter, pivot::FilterType::DateThisWeek,
                                     pivot::FilterType::DateM1,          pivot::FilterType::DateM12,
                                     pivot::FilterType::DateQ1,          pivot::FilterType::DateQ4};
  std::string filters;
  for (const char* name : names) {
    filters.append("<filter fld=\"2\" type=\"").append(name).append("\"><autoFilter ref=\"A1\"/></filter>");
  }
  const pivot::PivotTable table = Read(DefinitionWithFilters(filters, std::size(names)));
  ASSERT_EQ(table.active_filters().size(), std::size(names));
  for (std::size_t i = 0; i < std::size(names); ++i) {
    EXPECT_EQ(table.active_filters()[i].type, types[i]) << names[i];
    EXPECT_EQ(table.active_filters()[i].source->field_index, 2U) << names[i];
    EXPECT_TRUE(std::holds_alternative<std::monostate>(table.active_filters()[i].value)) << names[i];
  }
  const std::string round = write_pivot_table_definition(table);
  EXPECT_EQ(CountOccurrences(round, "<filter "), std::size(names)) << round;
}

TEST(PivotTableReader, MalformedCalendarSelectorsAreRetained) {
  // `M0`, `M13` and `Q5` are outside their families, and a bare letter or
  // a non-digit tail names nothing at all. None may decode to a range that
  // silently filters the wrong months.
  std::string filters;
  for (const char* type : {"M0", "M13", "Q0", "Q5", "M", "Mx"}) {
    filters.append("<filter fld=\"2\" type=\"").append(type).append("\"><autoFilter ref=\"A1\"/></filter>");
  }
  const pivot::PivotTable table = Read(DefinitionWithFilters(filters, 6U));
  EXPECT_TRUE(table.active_filters().empty());
  EXPECT_EQ(table.retained_filters().size(), 6U);
}

TEST(PivotTableReader, EveryValueAndDateComparisonTypeIsDecodedAndRoundTrips) {
  struct Case {
    const char* type;
    pivot::FilterType expected;
    bool ranged;
  };
  const Case cases[] = {
      {"valueEqual", pivot::FilterType::ValueEqual, false},
      {"valueNotEqual", pivot::FilterType::ValueNotEqual, false},
      {"valueGreaterThan", pivot::FilterType::ValueGreaterThan, false},
      {"valueGreaterThanOrEqual", pivot::FilterType::ValueGreaterThanOrEqual, false},
      {"valueLessThan", pivot::FilterType::ValueLessThan, false},
      {"valueLessThanOrEqual", pivot::FilterType::ValueLessThanOrEqual, false},
      {"valueBetween", pivot::FilterType::ValueBetween, true},
      {"valueNotBetween", pivot::FilterType::ValueNotBetween, true},
      {"dateEqual", pivot::FilterType::DateEqual, false},
      {"dateNotEqual", pivot::FilterType::DateNotEqual, false},
      {"dateOlderThan", pivot::FilterType::DateOlderThan, false},
      {"dateOlderThanOrEqual", pivot::FilterType::DateOlderThanOrEqual, false},
      {"dateNewerThan", pivot::FilterType::DateNewerThan, false},
      {"dateNewerThanOrEqual", pivot::FilterType::DateNewerThanOrEqual, false},
      {"dateBetween", pivot::FilterType::LabelDate, true},
      {"dateNotBetween", pivot::FilterType::DateNotBetween, true},
  };
  std::string filters;
  for (const Case& c : cases) {
    filters.append("<filter fld=\"0\" type=\"").append(c.type).append("\" iMeasureFld=\"0\"><autoFilter ref=\"A1\">");
    filters.append("<filterColumn colId=\"0\"><customFilters><customFilter operator=\"equal\" val=\"4\"/>");
    if (c.ranged) {
      filters.append("<customFilter operator=\"equal\" val=\"9\"/>");
    }
    filters.append("</customFilters></filterColumn></autoFilter></filter>");
  }
  const pivot::PivotTable table = Read(DefinitionWithFilters(filters, std::size(cases)));
  const auto& decoded = table.active_filters();
  ASSERT_EQ(decoded.size(), std::size(cases));
  for (std::size_t i = 0; i < decoded.size(); ++i) {
    EXPECT_EQ(decoded[i].type, cases[i].expected) << cases[i].type;
    EXPECT_EQ(decoded[i].value, Payload(4.0)) << cases[i].type;
    EXPECT_EQ(decoded[i].value_high, cases[i].ranged ? Payload(9.0) : Payload()) << cases[i].type;
  }
  const std::string round = write_pivot_table_definition(table);
  for (const Case& c : cases) {
    EXPECT_NE(round.find(std::string("type=\"") + c.type + "\""), std::string::npos) << c.type;
  }
}

TEST(PivotTableReader, ValueFilterCriteriaThatAreNotNumbersAreRetained) {
  // The criterion shares the sheet path's number lexer, so a caption-ish
  // payload under a value type is refused instead of reaching the
  // comparison as zero.
  const pivot::PivotTable table = Read(DefinitionWithFilters(
      "<filter fld=\"0\" type=\"valueGreaterThan\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><customFilters><customFilter operator=\"greaterThan\" val=\"North\"/>"
      "</customFilters></filterColumn></autoFilter></filter>",
      1U));
  EXPECT_TRUE(table.active_filters().empty());
  EXPECT_EQ(table.retained_filters().size(), 1U);
}

TEST(PivotTableReader, ValueBetweenWithOnlyOneBoundIsRetained) {
  const pivot::PivotTable table =
      Read(DefinitionWithFilters("<filter fld=\"0\" type=\"valueBetween\"><autoFilter ref=\"A1\">"
                                 "<filterColumn colId=\"0\"><customFilters and=\"1\">"
                                 "<customFilter operator=\"greaterThanOrEqual\" val=\"100\"/>"
                                 "</customFilters></filterColumn></autoFilter></filter>",
                                 1U));
  EXPECT_TRUE(table.active_filters().empty());
  EXPECT_EQ(table.retained_filters().size(), 1U);
}

}  // namespace
}  // namespace formulon::io
