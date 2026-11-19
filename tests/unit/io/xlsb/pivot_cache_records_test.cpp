#include "io/xlsb/pivot_cache_records.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "pivot_twin_corpus.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using cache_records::CellKind;
using cache_records::PivotCacheRecordsDoc;
using cache_records::RecordCell;
using cache_records::RecordDate;

constexpr CellKind kX = CellKind::kSharedIndex;
constexpr CellKind kN = CellKind::kNumber;
constexpr CellKind kS = CellKind::kText;
constexpr CellKind kD = CellKind::kDate;
constexpr CellKind kNE = CellKind::kNumberOrError;
constexpr CellKind kNB = CellKind::kNumberOrBlank;
constexpr CellKind kSB = CellKind::kTextOrBlank;
constexpr CellKind kDB = CellKind::kDateOrBlank;
constexpr CellKind kE = CellKind::kError;

/// BErr byte of #N/A.
constexpr std::uint8_t kErrNA = 0x2AU;

/// 2024-01-05, the first source date of the dated twins, as a 1900 and a 1904 serial.
constexpr double kJan5Serial1900 = 45296.0;
constexpr double kJan5Serial1904 = 45296.0 - 1462.0;

using Parts = std::vector<std::pair<std::string, std::vector<std::uint8_t>>>;

/// One records part with the definition of the same number beside it.
struct RecordsPart {
  std::string name;
  std::vector<std::uint8_t> records;
  std::vector<std::uint8_t> definition;
};

std::vector<RecordsPart> RecordsParts(const std::string& package_path) {
  const Parts records = pivot_twins::Parts(package_path, "xl/pivotCache/pivotCacheRecords");
  const Parts definitions = pivot_twins::Parts(package_path, "xl/pivotCache/pivotCacheDefinition");
  const std::string kRecordsPrefix = "xl/pivotCache/pivotCacheRecords";
  std::vector<RecordsPart> out;
  for (const auto& [name, bytes] : records) {
    const std::string definition_name = "xl/pivotCache/pivotCacheDefinition" + name.substr(kRecordsPrefix.size());
    RecordsPart part;
    part.name = name;
    part.records = bytes;
    for (const auto& [def_name, def_bytes] : definitions) {
      if (def_name == definition_name) {
        part.definition = def_bytes;
      }
    }
    EXPECT_FALSE(part.definition.empty()) << package_path << " " << name << " has no definition";
    out.push_back(std::move(part));
  }
  return out;
}

std::string Fixture(const std::string& stem) {
  for (const std::string& path : pivot_twins::CorpusPaths()) {
    const std::string suffix = "/" + stem + ".xlsb";
    if (path.size() >= suffix.size() && path.compare(path.size() - suffix.size(), suffix.size(), suffix) == 0) {
      return path;
    }
  }
  ADD_FAILURE() << "no corpus package " << stem;
  return {};
}

/// The first records part of `stem`, decoded against its own definition.
struct Decoded {
  std::vector<CellKind> kinds;
  PivotCacheRecordsDoc doc;
};

Decoded DecodeFirst(const std::string& stem) {
  Decoded out;
  const std::vector<RecordsPart> parts = RecordsParts(Fixture(stem));
  if (parts.empty()) {
    ADD_FAILURE() << stem << " has no records part";
    return out;
  }
  auto kinds = cell_kinds_of_cache_definition(SpanOf(parts.front().definition));
  if (!kinds) {
    ADD_FAILURE() << stem << ": " << kinds.error().message;
    return out;
  }
  out.kinds = kinds.value();
  auto doc = decode_pivot_cache_records(SpanOf(parts.front().records), out.kinds);
  if (!doc) {
    ADD_FAILURE() << stem << ": " << doc.error().message;
    return out;
  }
  out.doc = std::move(doc.value());
  return out;
}

void ExpectDate(const RecordCell& cell, std::uint16_t year, std::uint16_t month, std::uint8_t day) {
  ASSERT_EQ(cell.kind, kD);
  EXPECT_EQ(cell.date.year, year);
  EXPECT_EQ(cell.date.month, month);
  EXPECT_EQ(cell.date.day, day);
}

TEST(XlsbPivotCacheRecords, EveryCorpusRecordsPartRoundTripsByteForByte) {
  std::size_t parts_seen = 0;
  for (const std::string& path : pivot_twins::CorpusPaths()) {
    for (const RecordsPart& part : RecordsParts(path)) {
      ++parts_seen;
      auto kinds = cell_kinds_of_cache_definition(SpanOf(part.definition));
      ASSERT_TRUE(kinds) << path << " " << part.name << ": " << kinds.error().message;
      auto doc = decode_pivot_cache_records(SpanOf(part.records), kinds.value());
      ASSERT_TRUE(doc) << path << " " << part.name << ": " << doc.error().message;
      auto bytes = encode_pivot_cache_records(doc.value(), kinds.value());
      ASSERT_TRUE(bytes) << path << " " << part.name << ": " << bytes.error().message;
      EXPECT_EQ(bytes.value(), part.records) << path << " " << part.name;
    }
  }
  // 82 twins with one cache each, three caches in pivot_days_end, four in pivot_week_1900.
  EXPECT_EQ(parts_seen, 91U);
}

TEST(XlsbPivotCacheRecords, InlineDateTwinsDecodeAsIndexIndexDateNumberNumber) {
  const char* const kTwins[] = {
      "A8_custom_subtotal_countA",
      "A8_custom_subtotal_max",
      "A8_custom_subtotal_min",
      "A8_custom_subtotal_stdDev",
      "A8_custom_subtotal_stdDevP",
      "A8_custom_subtotal_varP",
      "B04_rowgrand_off",
      "B05_colgrand_off",
      "B07_data_position",
      "B14_merge_items",
      "B15_indent",
      "B17_page_field",
      "C01_subtotal_bottom",
      "C02_insert_blank_row",
      "C03_default_subtotal_off",
      "C11_item_collapsed",
      "D02_showas_running_total",
  };
  for (const char* twin : kTwins) {
    const Decoded decoded = DecodeFirst(twin);
    EXPECT_EQ(decoded.kinds, (std::vector<CellKind>{kX, kX, kD, kN, kN})) << twin;
    ASSERT_EQ(decoded.doc.rows.size(), 12U) << twin;
    // Twin .xlsx: <r><x v="0"/><x v="0"/><d v="2024-01-05T00:00:00"/><n v="10"/><n v="1"/></r>.
    const std::vector<RecordCell>& first = decoded.doc.rows.front();
    EXPECT_EQ(first[0].index, 0U) << twin;
    EXPECT_EQ(first[1].index, 0U) << twin;
    ExpectDate(first[2], 2024, 1, 5);
    EXPECT_EQ(first[3].number, 10.0) << twin;
    EXPECT_EQ(first[4].number, 1.0) << twin;
    auto serial = record_date_serial(first[2].date, /*date1904=*/false);
    ASSERT_TRUE(serial) << twin;
    EXPECT_EQ(serial.value(), kJan5Serial1900) << twin;
  }
}

TEST(XlsbPivotCacheRecords, MaxTwinMatchesItsXlsxRecords) {
  const Decoded decoded = DecodeFirst("A8_custom_subtotal_max");
  ASSERT_EQ(decoded.doc.rows.size(), 12U);
  // Twin .xlsx row 2: <x v="1"/><x v="0"/><d v="2024-02-06T00:00:00"/><n v="20"/><n v="2"/>.
  const std::vector<RecordCell>& second = decoded.doc.rows[1];
  EXPECT_EQ(second[0].index, 1U);
  EXPECT_EQ(second[1].index, 0U);
  ExpectDate(second[2], 2024, 2, 6);
  EXPECT_EQ(second[3].number, 20.0);
  EXPECT_EQ(second[4].number, 2.0);
  // Last row: x 2, x 1, 2024-12-16, 120, 4.
  const std::vector<RecordCell>& last = decoded.doc.rows.back();
  EXPECT_EQ(last[0].index, 2U);
  EXPECT_EQ(last[1].index, 1U);
  ExpectDate(last[2], 2024, 12, 16);
  EXPECT_EQ(last[3].number, 120.0);
  EXPECT_EQ(last[4].number, 4.0);
}

TEST(XlsbPivotCacheRecords, SumTwinIndexesFourFieldsAndInlinesTheFifth) {
  const Decoded decoded = DecodeFirst("A5_C9__sum");
  EXPECT_EQ(decoded.kinds, (std::vector<CellKind>{kX, kX, kX, kX, kN}));
  ASSERT_EQ(decoded.doc.rows.size(), 17U);
  // Twin .xlsx: <r><x v="0"/><x v="0"/><x v="0"/><x v="0"/><n v="0.3"/></r>.
  const std::vector<RecordCell>& first = decoded.doc.rows.front();
  for (std::size_t i = 0; i < 4; ++i) {
    EXPECT_EQ(first[i].index, 0U) << i;
  }
  EXPECT_EQ(first[4].number, 0.3);
  // <r><x v="0"/><x v="1"/><x v="0"/><x v="0"/><n v="0.7"/></r>.
  EXPECT_EQ(decoded.doc.rows[1][1].index, 1U);
  EXPECT_EQ(decoded.doc.rows[1][4].number, 0.7);
}

TEST(XlsbPivotCacheRecords, CalcItemTwinCarriesAnInlineString) {
  const Decoded decoded = DecodeFirst("C13_calc_item");
  EXPECT_EQ(decoded.kinds, (std::vector<CellKind>{kX, kS, kD, kN, kN}));
  ASSERT_EQ(decoded.doc.rows.size(), 12U);
  // Twin .xlsx: <r><x v="0"/><s v="Apple"/><d v="2024-01-05T00:00:00"/><n v="10"/><n v="1"/></r>.
  const std::vector<RecordCell>& first = decoded.doc.rows.front();
  EXPECT_EQ(first[0].index, 0U);
  EXPECT_EQ(first[1].kind, kS);
  EXPECT_EQ(first[1].text, "Apple");
  ExpectDate(first[2], 2024, 1, 5);
  EXPECT_EQ(first[3].number, 10.0);
  EXPECT_EQ(first[4].number, 1.0);
  // Last row: x 2, s Pear, 2024-12-16, 120, 4.
  const std::vector<RecordCell>& last = decoded.doc.rows.back();
  EXPECT_EQ(last[1].text, "Pear");
  ExpectDate(last[2], 2024, 12, 16);
}

TEST(XlsbPivotCacheRecords, GroupingBooksSkipTheGroupOnlyField) {
  for (const char* stem : {"pivot_days_end", "pivot_week_1900"}) {
    for (const RecordsPart& part : RecordsParts(Fixture(stem))) {
      // Fields d, v, and the grouping-only `日 (d)`, which has no record cell.
      auto kinds = cell_kinds_of_cache_definition(SpanOf(part.definition));
      ASSERT_TRUE(kinds) << stem << " " << part.name;
      EXPECT_EQ(kinds.value(), (std::vector<CellKind>{kX, kN})) << stem << " " << part.name;
      auto doc = decode_pivot_cache_records(SpanOf(part.records), kinds.value());
      ASSERT_TRUE(doc) << stem << " " << part.name;
      ASSERT_EQ(doc.value().rows.size(), 71U) << stem << " " << part.name;
      EXPECT_EQ(doc.value().rows.front()[0].index, 0U);
      EXPECT_EQ(doc.value().rows.front()[1].number, 1.0);
      EXPECT_EQ(doc.value().rows.back()[0].index, 70U);
    }
  }
}

TEST(XlsbPivotCacheRecords, RefusesKindsThatDoNotExplainTheRowWidth) {
  const std::vector<RecordsPart> parts = RecordsParts(Fixture("C13_calc_item"));
  ASSERT_FALSE(parts.empty());
  const ByteSpan records = SpanOf(parts.front().records);
  // The text cell read as an index leaves bytes over; the date read as a number too, but one field short.
  for (const std::vector<CellKind>& wrong :
       {std::vector<CellKind>{kX, kX, kD, kN, kN}, std::vector<CellKind>{kX, kS, kN, kN},
        std::vector<CellKind>{kX, kS, kD, kN, kN, kN}, std::vector<CellKind>{}}) {
    auto doc = decode_pivot_cache_records(records, wrong);
    ASSERT_FALSE(doc);
    EXPECT_EQ(doc.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
  }
}

TEST(XlsbPivotCacheRecords, RefusesAnHourPastTheDay) {
  const std::vector<RecordsPart> parts = RecordsParts(Fixture("A8_custom_subtotal_max"));
  ASSERT_FALSE(parts.empty());
  std::vector<std::uint8_t> bytes = parts.front().records;
  const std::vector<FramedRecord> recs = split_records(bytes);
  ASSERT_GE(recs.size(), 2U);
  ASSERT_EQ(recs[1].type, pivot_records::kPCRRecord);
  // Row x, x, d: the date's hour byte is payload offset 4 + 4 + 5.
  const auto hour_at = static_cast<std::size_t>(recs[1].payload.data - bytes.data()) + 13U;
  bytes[hour_at] = 24U;
  auto doc = decode_pivot_cache_records(SpanOf(bytes), {kX, kX, kD, kN, kN});
  ASSERT_FALSE(doc);
  EXPECT_EQ(doc.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}

TEST(XlsbPivotCacheRecords, RefusesAnUnknownRecordAndAWrongCount) {
  std::vector<std::uint8_t> payload;
  std::vector<std::uint8_t> part;
  emit_u32(payload, 1U);
  emit_record(part, pivot_records::kBeginPivotCacheRecords, payload);
  payload.clear();
  emit_double(payload, 1.0);
  emit_record(part, pivot_records::kPCRRecord, payload);
  std::vector<std::uint8_t> with_unknown = part;
  emit_record(with_unknown, 35, std::vector<std::uint8_t>{});
  emit_record(part, pivot_records::kEndPivotCacheRecords, std::vector<std::uint8_t>{});
  emit_record(with_unknown, pivot_records::kEndPivotCacheRecords, std::vector<std::uint8_t>{});

  auto ok = decode_pivot_cache_records(SpanOf(part), {kN});
  ASSERT_TRUE(ok);
  ASSERT_EQ(ok.value().rows.size(), 1U);
  EXPECT_EQ(ok.value().rows[0][0].number, 1.0);

  auto unknown = decode_pivot_cache_records(SpanOf(with_unknown), {kN});
  ASSERT_FALSE(unknown);
  EXPECT_EQ(unknown.error().code, FormulonErrorCode::kIoXlsbUnknownRecord);

  std::vector<std::uint8_t> wrong_count = part;
  const std::vector<FramedRecord> recs = split_records(wrong_count);
  wrong_count[static_cast<std::size_t>(recs[0].payload.data - wrong_count.data())] = 2U;
  auto counted = decode_pivot_cache_records(SpanOf(wrong_count), {kN});
  ASSERT_FALSE(counted);
  EXPECT_EQ(counted.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}

/// The two Mac saves whose number field holds one inline #N/A.
std::vector<std::string> InlineErrorPaths() {
  const std::string root = std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_show_as/";
  return {root + "S8e_pctdiffPrevR__sum.xlsb", root + "S8e_runR__sum.xlsb"};
}

/// A records part of one row: BrtPCRRecordDt, then `items` as (record type, payload).
std::vector<std::uint8_t> DynamicRowPart(
    const std::vector<std::pair<std::uint16_t, std::vector<std::uint8_t>>>& items) {
  std::vector<std::uint8_t> payload;
  std::vector<std::uint8_t> part;
  emit_u32(payload, 1U);
  emit_record(part, pivot_records::kBeginPivotCacheRecords, payload);
  emit_record(part, pivot_records::kPCRRecordDt, std::vector<std::uint8_t>{});
  for (const auto& [type, bytes] : items) {
    emit_record(part, type, bytes);
  }
  emit_record(part, pivot_records::kEndPivotCacheRecords, std::vector<std::uint8_t>{});
  return part;
}

TEST(XlsbPivotCacheRecords, InlineErrorRowTakesTheDynamicForm) {
  for (const std::string& path : InlineErrorPaths()) {
    const std::vector<RecordsPart> parts = RecordsParts(path);
    ASSERT_EQ(parts.size(), 1U) << path;
    auto kinds = cell_kinds_of_cache_definition(SpanOf(parts.front().definition));
    ASSERT_TRUE(kinds) << path << ": " << kinds.error().message;
    EXPECT_EQ(kinds.value(), (std::vector<CellKind>{kX, kX, kNE})) << path;
    auto doc = decode_pivot_cache_records(SpanOf(parts.front().records), kinds.value());
    ASSERT_TRUE(doc) << path << ": " << doc.error().message;
    const std::vector<std::vector<RecordCell>>& rows = doc.value().rows;
    ASSERT_EQ(rows.size(), 13U) << path;
    // Row 11 is x 3, x 2, #N/A; every other Amt cell is a number.
    for (std::size_t i = 0; i < rows.size(); ++i) {
      EXPECT_EQ(rows[i][2].kind, i == 11U ? kE : kN) << path << " row " << i;
    }
    EXPECT_EQ(rows[11][0].index, 3U);
    EXPECT_EQ(rows[11][1].index, 2U);
    EXPECT_EQ(rows[11][2].error, kErrNA);
    EXPECT_EQ(rows[0][2].number, 1e16);
    EXPECT_EQ(rows[12][2].number, 7.0);
    auto bytes = encode_pivot_cache_records(doc.value(), kinds.value());
    ASSERT_TRUE(bytes) << path << ": " << bytes.error().message;
    EXPECT_EQ(bytes.value(), parts.front().records) << path;
  }
}

TEST(XlsbPivotCacheRecords, DynamicRowRoundTripsByteForByte) {
  std::vector<std::uint8_t> index;
  emit_u32(index, 5U);
  const std::vector<std::uint8_t> part =
      DynamicRowPart({{pivot_records::kPCDIIndex, index}, {pivot_records::kPCDIError, {kErrNA}}});
  auto doc = decode_pivot_cache_records(SpanOf(part), {kX, kNE});
  ASSERT_TRUE(doc) << doc.error().message;
  ASSERT_EQ(doc.value().rows.size(), 1U);
  EXPECT_EQ(doc.value().rows[0][0].index, 5U);
  EXPECT_EQ(doc.value().rows[0][1].kind, kE);
  EXPECT_EQ(doc.value().rows[0][1].error, kErrNA);
  auto bytes = encode_pivot_cache_records(doc.value(), {kX, kNE});
  ASSERT_TRUE(bytes) << bytes.error().message;
  EXPECT_EQ(bytes.value(), part);
}

TEST(XlsbPivotCacheRecords, RefusesUnmeasuredDynamicRows) {
  std::vector<std::uint8_t> number;
  emit_double(number, 1.0);
  std::vector<std::uint8_t> text;
  emit_xlwidestring(text, "a");
  std::vector<std::uint8_t> index;
  emit_u32(index, 0U);
  // No error or blank in the row, an error in a plain number field, a blank in a field that admits none, and a
  // text item for a number field.
  const std::vector<std::pair<std::vector<CellKind>, std::vector<std::uint8_t>>> cases = {
      {{kX, kNE}, DynamicRowPart({{pivot_records::kPCDIIndex, index}, {pivot_records::kPCDINumber, number}})},
      {{kX, kN}, DynamicRowPart({{pivot_records::kPCDIIndex, index}, {pivot_records::kPCDIError, {kErrNA}}})},
      {{kX, kN}, DynamicRowPart({{pivot_records::kPCDIIndex, index}, {pivot_records::kPCDIMissing, {}}})},
      {{kS, kNB}, DynamicRowPart({{pivot_records::kPCDIString, text}, {pivot_records::kPCDIString, text}})},
      {{kX, kNE}, DynamicRowPart({{pivot_records::kPCDIIndex, index}, {pivot_records::kPCDIIndex, index}})},
  };
  for (std::size_t i = 0; i < cases.size(); ++i) {
    auto doc = decode_pivot_cache_records(SpanOf(cases[i].second), cases[i].first);
    ASSERT_FALSE(doc) << i;
    EXPECT_EQ(doc.error().code, FormulonErrorCode::kIoXlsbUnknownRecord) << i;
  }

  // An error outside a number-or-error field.
  RecordCell error;
  error.kind = kE;
  error.error = kErrNA;
  PivotCacheRecordsDoc plain_number;
  plain_number.rows.push_back({error});
  EXPECT_FALSE(encode_pivot_cache_records(plain_number, {kN}));
  EXPECT_FALSE(encode_pivot_cache_records(plain_number, {kE}));
}

// A dynamic row spells each cell by its own item record, so a number or a text cell beside an error round trips.
TEST(XlsbPivotCacheRecords, DynamicRowSpellsEveryCellKind) {
  std::vector<std::uint8_t> number;
  emit_double(number, 1.5);
  std::vector<std::uint8_t> text;
  emit_xlwidestring(text, "Fig");
  std::vector<std::uint8_t> date = {0xE8, 0x07, 0x02, 0x00, 0x0F, 0x0C, 0x00, 0x00};
  const std::vector<std::uint8_t> part = DynamicRowPart({{pivot_records::kPCDIDatetime, date},
                                                         {pivot_records::kPCDIString, text},
                                                         {pivot_records::kPCDINumber, number},
                                                         {pivot_records::kPCDIError, {kErrNA}},
                                                         {pivot_records::kPCDIMissing, {}}});
  const std::vector<CellKind> kinds = {kD, kS, kN, kNE, kNB};
  auto doc = decode_pivot_cache_records(SpanOf(part), kinds);
  ASSERT_TRUE(doc) << doc.error().message;
  const std::vector<RecordCell>& row = doc.value().rows.at(0);
  ExpectDate(row[0], 2024, 2, 15);
  EXPECT_EQ(row[1].text, "Fig");
  EXPECT_EQ(row[2].number, 1.5);
  EXPECT_EQ(row[3].kind, kE);
  EXPECT_EQ(row[4].kind, CellKind::kMissing);
  auto bytes = encode_pivot_cache_records(doc.value(), kinds);
  ASSERT_TRUE(bytes) << bytes.error().message;
  EXPECT_EQ(bytes.value(), part);
}

// Excel's saves of a source row left blank (xlsb_parity/pivot/blank): a field with shared items indexes its missing
// item in a fixed row; a field without them makes that one row dynamic, with BrtPCDIMissing for the blank cell.
TEST(XlsbPivotCacheRecords, BlankSourceCellsDecodeAndRoundTrip) {
  const std::string root = std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/pivot/blank/";
  constexpr CellKind kM = CellKind::kMissing;
  const struct {
    const char* stem;
    std::vector<CellKind> kinds;
    std::vector<CellKind> blank_row;  ///< The cells of the one row holding a blank; empty when no row does.
  } kCases[] = {
      {"blank_control", {kD, kX, kS, kN, kN}, {}},
      {"blank_value_only", {kD, kX, kS, kNB, kN}, {kD, kX, kS, kM, kN}},
      {"blank_date_not_on_axis", {kDB, kX, kS, kN, kN}, {kM, kX, kS, kN, kN}},
      {"blank_whole_row", {kDB, kX, kSB, kNB, kNB}, {kM, kX, kM, kM, kM}},
      {"blank_row_field_only", {kD, kX, kS, kN, kN}, {}},
  };
  for (const auto& c : kCases) {
    SCOPED_TRACE(c.stem);
    const std::vector<RecordsPart> parts = RecordsParts(root + c.stem + ".xlsb");
    ASSERT_EQ(parts.size(), 1U);
    auto kinds = cell_kinds_of_cache_definition(SpanOf(parts.front().definition));
    ASSERT_TRUE(kinds) << kinds.error().message;
    EXPECT_EQ(kinds.value(), c.kinds);
    auto doc = decode_pivot_cache_records(SpanOf(parts.front().records), kinds.value());
    ASSERT_TRUE(doc) << doc.error().message;
    ASSERT_EQ(doc.value().rows.size(), 13U);
    std::vector<std::vector<CellKind>> blank_rows;
    for (const std::vector<RecordCell>& row : doc.value().rows) {
      std::vector<CellKind> cells;
      bool blank = false;
      for (const RecordCell& cell : row) {
        cells.push_back(cell.kind);
        blank = blank || cell.kind == kM;
      }
      if (blank) {
        blank_rows.push_back(cells);
      }
    }
    EXPECT_EQ(blank_rows, c.blank_row.empty() ? std::vector<std::vector<CellKind>>{}
                                              : std::vector<std::vector<CellKind>>{c.blank_row});
    auto bytes = encode_pivot_cache_records(doc.value(), kinds.value());
    ASSERT_TRUE(bytes) << bytes.error().message;
    EXPECT_EQ(bytes.value(), parts.front().records);
  }
}

TEST(XlsbPivotCacheRecords, EncodeRefusesACellOfTheWrongKind) {
  PivotCacheRecordsDoc doc;
  RecordCell cell;
  cell.kind = kN;
  cell.number = 1.0;
  doc.rows.push_back({cell});
  EXPECT_TRUE(encode_pivot_cache_records(doc, {kN}));
  EXPECT_FALSE(encode_pivot_cache_records(doc, {kX}));
  EXPECT_FALSE(encode_pivot_cache_records(doc, {kN, kN}));
}

TEST(XlsbPivotCacheRecords, DateCellsConvertToSerialsBothWays) {
  const RecordDate jan5{2024, 1, 5};
  auto serial1900 = record_date_serial(jan5, /*date1904=*/false);
  ASSERT_TRUE(serial1900);
  EXPECT_EQ(serial1900.value(), kJan5Serial1900);
  auto serial1904 = record_date_serial(jan5, /*date1904=*/true);
  ASSERT_TRUE(serial1904);
  EXPECT_EQ(serial1904.value(), kJan5Serial1904);

  for (const bool date1904 : {false, true}) {
    auto back = record_date_from_serial(date1904 ? kJan5Serial1904 : kJan5Serial1900, date1904);
    ASSERT_TRUE(back) << date1904;
    EXPECT_EQ(back.value().year, 2024U);
    EXPECT_EQ(back.value().month, 1U);
    EXPECT_EQ(back.value().day, 5U);
  }
  // 1900-03-01 is the first 1900-system day; 1904-01-01 is 1904 serial 0.
  auto march1 = record_date_from_serial(61.0, /*date1904=*/false);
  ASSERT_TRUE(march1);
  EXPECT_EQ(march1.value().month, 3U);
  EXPECT_EQ(march1.value().day, 1U);
  auto epoch1904 = record_date_from_serial(0.0, /*date1904=*/true);
  ASSERT_TRUE(epoch1904);
  EXPECT_EQ(epoch1904.value().year, 1904U);

  auto noon = record_date_from_serial(kJan5Serial1900 + 0.5, /*date1904=*/false);
  ASSERT_TRUE(noon);
  EXPECT_EQ(noon.value().hour, 12U);
  EXPECT_EQ(noon.value().minute, 0U);
  EXPECT_FALSE(record_date_from_serial(kJan5Serial1900 + 0.5 / 86400.0, /*date1904=*/false));
  EXPECT_FALSE(record_date_from_serial(60.0, /*date1904=*/false));
  EXPECT_FALSE(record_date_serial(RecordDate{1900, 2, 28}, /*date1904=*/false));
  EXPECT_FALSE(record_date_serial(RecordDate{2023, 2, 29}, /*date1904=*/false));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
