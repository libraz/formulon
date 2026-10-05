// Pivot-table definition reader tests grouped by XML concern.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
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
TEST(PivotTableReader, AuthoredFiltersDecodeForEvaluationAndStillRoundTripAsPassthrough) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("<pivotTableStyleInfo name=\"PivotStyleLight16\"/>");
  xml.append(
      "<filters count=\"1\"><filter fld=\"0\" type=\"captionEqual\" evalOrder=\"-1\" id=\"1\">"
      "<autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"equal\" val=\"North\"/></customFilters>"
      "</filterColumn></autoFilter></filter></filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  EXPECT_TRUE(table.active_filters().empty());
  ASSERT_EQ(table.authored_caption_filters().size(), 1U);
  EXPECT_EQ(table.authored_caption_filters()[0].field_index, 0U);
  EXPECT_EQ(table.authored_caption_filters()[0].predicate, pivot::CaptionPredicate::Equal);
  EXPECT_EQ(table.authored_caption_filters()[0].value, "North");
  EXPECT_NE(table.raw_passthrough_xml().find("<filters"), std::string::npos);
  EXPECT_NE(table.raw_passthrough_xml().find("val=\"North\""), std::string::npos);

  const std::string round = write_pivot_table_definition(table);
  EXPECT_EQ(CountOccurrences(round, "<filters"), 1U) << round;
  EXPECT_NE(round.find("val=\"North\""), std::string::npos) << round;
  // The style block is authored before `<filters>`; re-emission keeps that
  // order, which the schema requires.
  EXPECT_LT(round.find("<pivotTableStyleInfo"), round.find("<filters"));

  auto reparsed_or = read_pivot_table_definition(Bytes(round));
  ASSERT_TRUE(static_cast<bool>(reparsed_or)) << reparsed_or.error().message;
  EXPECT_TRUE(reparsed_or.value().active_filters().empty());
  // The decode is stable across the round trip, and the writer still emits
  // the block from the passthrough tail rather than from the decoded view.
  EXPECT_EQ(reparsed_or.value().authored_caption_filters().size(), 1U);
  EXPECT_EQ(write_pivot_table_definition(reparsed_or.value()), round);
}
TEST(PivotTableReader, AuthoredCaptionFilterVariantsDecodeAndNonCaptionTypesAreSkipped) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"5\">");
  xml.append(
      "<filter fld=\"1\" type=\"captionBeginsWith\"><autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"equal\" val=\"No*\"/></customFilters>"
      "</filterColumn></autoFilter></filter>");
  xml.append(
      "<filter fld=\"2\" type=\"captionContains\"><autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"equal\" val=\"*or*\"/></customFilters>"
      "</filterColumn></autoFilter></filter>");
  xml.append(
      "<filter fld=\"3\" type=\"captionEndsWith\"><autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"equal\" val=\"*th\"/></customFilters>"
      "</filterColumn></autoFilter></filter>");
  xml.append(
      "<filter fld=\"4\" type=\"captionBetween\"><autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters and=\"1\"><customFilter operator=\"greaterThanOrEqual\" val=\"B\"/>"
      "<customFilter operator=\"lessThanOrEqual\" val=\"M\"/></customFilters>"
      "</filterColumn></autoFilter></filter>");
  // Value family: outside the caption set, so passthrough-only.
  xml.append(
      "<filter fld=\"5\" type=\"valueGreaterThan\" iMeasureFld=\"0\"><autoFilter ref=\"A3:A9\">"
      "<filterColumn colId=\"0\"><customFilters><customFilter operator=\"greaterThan\" val=\"100\"/>"
      "</customFilters></filterColumn></autoFilter></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const auto& decoded = table_or.value().authored_caption_filters();
  ASSERT_EQ(decoded.size(), 4U);
  EXPECT_EQ(decoded[0].field_index, 1U);
  EXPECT_EQ(decoded[0].predicate, pivot::CaptionPredicate::BeginsWith);
  EXPECT_EQ(decoded[0].value, "No");
  EXPECT_EQ(decoded[1].predicate, pivot::CaptionPredicate::Contains);
  EXPECT_EQ(decoded[1].value, "or");
  EXPECT_EQ(decoded[2].predicate, pivot::CaptionPredicate::EndsWith);
  EXPECT_EQ(decoded[2].value, "th");
  EXPECT_EQ(decoded[3].predicate, pivot::CaptionPredicate::Between);
  EXPECT_EQ(decoded[3].value, "B");
  EXPECT_EQ(decoded[3].value_high, "M");
  // Every entry, decoded or not, is still re-emitted from the passthrough.
  const std::string round = write_pivot_table_definition(table_or.value());
  EXPECT_NE(round.find("valueGreaterThan"), std::string::npos) << round;
  EXPECT_EQ(CountOccurrences(round, "<filter "), 5U) << round;
}
TEST(PivotTableReader, AuthoredCaptionFilterAcceptsThePlainFilterListShape) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append(
      "<filters count=\"1\"><filter fld=\"0\" type=\"captionEqual\"><autoFilter ref=\"A3:A9\">"
      "<filterColumn colId=\"0\"><filters><filter val=\"South\"/></filters>"
      "</filterColumn></autoFilter></filter></filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  ASSERT_EQ(table_or.value().authored_caption_filters().size(), 1U);
  EXPECT_EQ(table_or.value().authored_caption_filters()[0].value, "South");
}
TEST(PivotTableReader, UnevaluableFilterFamiliesAreSkippedButStillRoundTrip) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"3\">");
  // A count filter with no nested `<top10>` names no quantity at all.
  xml.append("<filter fld=\"0\" type=\"count\"><autoFilter ref=\"A1\"/></filter>");
  // An unknown type is the forward-compatibility case: a family this
  // version does not know must be passed through rather than guessed at.
  xml.append("<filter fld=\"1\" type=\"someFutureFamily\"><autoFilter ref=\"A1\"/></filter>");
  // A caption filter whose criterion element is missing entirely.
  xml.append("<filter fld=\"2\" type=\"captionEqual\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_TRUE(table_or.value().authored_value_filters().empty());
  EXPECT_TRUE(table_or.value().authored_caption_filters().empty());
  EXPECT_TRUE(table_or.value().authored_period_filters().empty());
  EXPECT_TRUE(table_or.value().authored_recurring_filters().empty());

  const std::string round = write_pivot_table_definition(table_or.value());
  EXPECT_EQ(CountOccurrences(round, "<filter "), 3U) << round;
  EXPECT_NE(round.find("someFutureFamily"), std::string::npos) << round;
}
TEST(PivotTableReader, DecodedFilterFamiliesStillRoundTripVerbatim) {
  // Decoding is read-side only: the writer re-emits the `<filters>` block
  // from the verbatim passthrough tail, so gaining a decoder must not
  // change a single byte of what is written back.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"2\">");
  xml.append("<filter fld=\"2\" type=\"thisWeek\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"2\" type=\"M1\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_EQ(table_or.value().authored_period_filters().size(), 1U);
  EXPECT_EQ(table_or.value().authored_recurring_filters().size(), 1U);

  const std::string round = write_pivot_table_definition(table_or.value());
  EXPECT_EQ(CountOccurrences(round, "<filter "), 2U) << round;
  EXPECT_NE(round.find("thisWeek"), std::string::npos) << round;
  EXPECT_NE(round.find("\"M1\""), std::string::npos) << round;
}
TEST(PivotTableReader, TopNFlavoursAreDistinguishedByTheirTypeAlone) {
  // The three flavours share an identically shaped `<top10 val="N">`, so a
  // decode that ignored the type attribute would silently conflate them.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"3\">");
  xml.append(
      "<filter fld=\"0\" type=\"count\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<top10 val=\"2\"/></filterColumn></autoFilter></filter>");
  xml.append(
      "<filter fld=\"0\" type=\"percent\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<top10 percent=\"1\" val=\"70\"/></filterColumn></autoFilter></filter>");
  xml.append(
      "<filter fld=\"0\" type=\"sum\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<top10 val=\"100\"/></filterColumn></autoFilter></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const auto& filters = table_or.value().authored_value_filters();
  ASSERT_EQ(filters.size(), 3U);
  for (const auto& entry : filters) {
    EXPECT_EQ(entry.type, pivot::FilterType::ValueTop10);
  }
  EXPECT_EQ(filters[0].top_n_basis, pivot::TopNBasis::Items);
  EXPECT_DOUBLE_EQ(filters[0].value, 2.0);
  EXPECT_EQ(filters[1].top_n_basis, pivot::TopNBasis::Percent);
  EXPECT_DOUBLE_EQ(filters[1].value, 70.0);
  EXPECT_EQ(filters[2].top_n_basis, pivot::TopNBasis::Sum);
  EXPECT_DOUBLE_EQ(filters[2].value, 100.0);
}
TEST(PivotTableReader, TopTenDirectionDecodesTopAttributeAndDefaultsToTop) {
  // Excel's "Bottom N" dialog option writes `top="0"` on the same
  // `<top10>` element as "Top N"; the attribute is absent for "Top N"
  // itself, so a missing attribute must decode the same as `top="1"`.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"2\">");
  xml.append(
      "<filter fld=\"0\" type=\"count\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<top10 val=\"2\"/></filterColumn></autoFilter></filter>");
  xml.append(
      "<filter fld=\"0\" type=\"count\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<top10 top=\"0\" val=\"2\"/></filterColumn></autoFilter></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const auto& filters = table_or.value().authored_value_filters();
  ASSERT_EQ(filters.size(), 2U);
  EXPECT_TRUE(filters[0].top);
  EXPECT_FALSE(filters[1].top);
}
TEST(PivotTableReader, RelativePeriodFiltersDecodeFromTheTypeNameAlone) {
  // These carry no criteria: the window is implied by the type, so the
  // decode is the name mapping and nothing else.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"3\">");
  xml.append("<filter fld=\"2\" type=\"thisMonth\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"3\" type=\"yearToDate\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"1\" type=\"lastQuarter\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const auto& periods = table_or.value().authored_period_filters();
  ASSERT_EQ(periods.size(), 3U);
  EXPECT_EQ(periods[0].field_index, 2U);
  EXPECT_EQ(periods[0].period, pivot::RelativePeriod::ThisMonth);
  EXPECT_EQ(periods[1].field_index, 3U);
  EXPECT_EQ(periods[1].period, pivot::RelativePeriod::YearToDate);
  EXPECT_EQ(periods[2].field_index, 1U);
  EXPECT_EQ(periods[2].period, pivot::RelativePeriod::LastQuarter);
  // A period entry carries no bounds, so it must not land in the list whose
  // members all do.
  EXPECT_TRUE(table_or.value().authored_value_filters().empty());

  const std::string round = write_pivot_table_definition(table_or.value());
  EXPECT_EQ(CountOccurrences(round, "<filter "), 3U) << round;
}
TEST(PivotTableReader, RecurringPeriodFiltersDecodeToACalendarMonthRange) {
  // `M<n>` is one month and `Q<n>` its three, both year-free -- so unlike
  // the relative periods these need no clock to mean something.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"4\">");
  xml.append("<filter fld=\"2\" type=\"M1\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"2\" type=\"M12\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"3\" type=\"Q1\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"1\" type=\"Q4\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const auto& recurring = table_or.value().authored_recurring_filters();
  ASSERT_EQ(recurring.size(), 4U);
  EXPECT_EQ(recurring[0].field_index, 2U);
  EXPECT_EQ(recurring[0].month_low, 1U);
  EXPECT_EQ(recurring[0].month_high, 1U);
  EXPECT_EQ(recurring[1].month_low, 12U);
  EXPECT_EQ(recurring[1].month_high, 12U);
  EXPECT_EQ(recurring[2].field_index, 3U);
  EXPECT_EQ(recurring[2].month_low, 1U);
  EXPECT_EQ(recurring[2].month_high, 3U);
  EXPECT_EQ(recurring[3].month_low, 10U);
  EXPECT_EQ(recurring[3].month_high, 12U);
  // A recurring entry is not a window and carries no bounds, so it must
  // reach neither sibling list.
  EXPECT_TRUE(table_or.value().authored_period_filters().empty());
  EXPECT_TRUE(table_or.value().authored_value_filters().empty());
}
TEST(PivotTableReader, MalformedRecurringSelectorsAreSkipped) {
  // `M0`, `M13` and `Q5` are outside their families, and a bare letter or
  // a non-digit tail names nothing at all. None may decode to a range that
  // silently filters the wrong months.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"6\">");
  for (const char* type : {"M0", "M13", "Q0", "Q5", "M", "Mx"}) {
    xml.append("<filter fld=\"2\" type=\"").append(type).append("\"><autoFilter ref=\"A1\"/></filter>");
  }
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_TRUE(table_or.value().authored_recurring_filters().empty());
}
TEST(PivotTableReader, WeekPeriodFiltersDecodeAlongsideTheOtherWindows) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"3\">");
  xml.append("<filter fld=\"2\" type=\"thisWeek\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"2\" type=\"lastWeek\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"2\" type=\"nextWeek\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const auto& periods = table_or.value().authored_period_filters();
  ASSERT_EQ(periods.size(), 3U);
  EXPECT_EQ(periods[0].period, pivot::RelativePeriod::ThisWeek);
  EXPECT_EQ(periods[1].period, pivot::RelativePeriod::LastWeek);
  EXPECT_EQ(periods[2].period, pivot::RelativePeriod::NextWeek);
}
TEST(PivotTableReader, ValueFilterCriteriaThatAreNotNumbersAreSkipped) {
  // The criterion shares the sheet path's number lexer, so a caption-ish
  // payload under a value type is refused instead of reaching the
  // comparison as zero.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append(
      "<filters count=\"1\"><filter fld=\"0\" type=\"valueGreaterThan\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><customFilters><customFilter operator=\"greaterThan\" val=\"North\"/>"
      "</customFilters></filterColumn></autoFilter></filter></filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_TRUE(table_or.value().authored_value_filters().empty());
}
TEST(PivotTableReader, ValueBetweenWithOnlyOneBoundIsSkipped) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append(
      "<filters count=\"1\"><filter fld=\"0\" type=\"valueBetween\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><customFilters and=\"1\">"
      "<customFilter operator=\"greaterThanOrEqual\" val=\"100\"/>"
      "</customFilters></filterColumn></autoFilter></filter></filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_TRUE(table_or.value().authored_value_filters().empty());
}

}  // namespace
}  // namespace formulon::io
