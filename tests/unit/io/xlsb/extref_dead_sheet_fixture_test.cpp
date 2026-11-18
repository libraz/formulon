//
// A reference into a sheet deleted from another workbook, against the Excel-saved
// `xlsb_extref_dead_{sheet,span,err3d,name}_B.xlsb` and `xlsb_extref_middle_deleted_B.xlsb`
// (column A), each beside its `.xlsx` twin.
//
// Excel stores such a reference as a `PtgRef3d` / `PtgArea3d` (or the error forms
// `PtgRefErr3d` / `PtgAreaErr3d`) whose XTI is (external book, itabFirst -1, itabLast -1); the
// xlsx twin stores bare `#REF!`. Tokens are compared per cell as (decoded XTI, rest of the
// rgce): the XTI table may shrink when two entries are the same.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <map>
#include <string>
#include <tuple>
#include <utility>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

struct FixtureSpec {
  const char* stem;
  std::uint32_t rows;
  /// Rows (0-based) whose formula reads a deleted sheet of the external book.
  std::vector<std::uint32_t> dead_rows;
  /// Defined names (all spelled `#REF!`) the workbook carries.
  std::uint32_t dead_names;
  /// Rows (0-based) of a legacy array formula, whose operand class this change does not touch.
  std::vector<std::uint32_t> array_rows;
};

const FixtureSpec kFixtures[] = {
    {"xlsb_extref_dead_sheet_B", 9, {0, 2, 3, 4}, 0, {7}}, {"xlsb_extref_dead_span_B", 9, {1, 5, 6, 7, 8}, 0, {}},
    {"xlsb_extref_dead_err3d_B", 3, {0, 1}, 0, {}},        {"xlsb_extref_dead_name_B", 2, {}, 1, {}},
    {"xlsb_extref_middle_deleted_B", 3, {}, 0, {}},
};

std::string FixturePath(const FixtureSpec& spec, const char* ext) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + spec.stem + "." + ext;
}

bool IsDead(const FixtureSpec& spec, std::uint32_t row) {
  return std::find(spec.dead_rows.begin(), spec.dead_rows.end(), row) != spec.dead_rows.end();
}

std::uint32_t U32(ByteSpan span, std::size_t at) {
  return static_cast<std::uint32_t>(span.data[at]) | (static_cast<std::uint32_t>(span.data[at + 1]) << 8) |
         (static_cast<std::uint32_t>(span.data[at + 2]) << 16) | (static_cast<std::uint32_t>(span.data[at + 3]) << 24);
}

/// One XTI with its book as 0 (this workbook) or the 1-based order among the external books.
using Xti = std::tuple<std::uint32_t, std::int32_t, std::int32_t>;

std::vector<Xti> ReadXtiTable(const std::vector<std::uint8_t>& package) {
  std::vector<Xti> out;
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  auto workbook = zip.read_entry("xl/workbook.bin");
  if (!workbook) {
    ADD_FAILURE() << "workbook.bin missing";
    return out;
  }
  std::vector<std::uint32_t> book_of_sup;
  std::uint32_t externals = 0;
  ByteSpan cursor{workbook.value().data(), workbook.value().size()};
  while (cursor.size > 0U) {
    auto record = read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "unreadable record";
      break;
    }
    const XlsbRecord& r = record.value();
    switch (static_cast<XlsbRecordType>(r.type)) {
      case XlsbRecordType::BrtSupSelf:
        book_of_sup.push_back(0U);
        break;
      case XlsbRecordType::BrtSupBookSrc:
        book_of_sup.push_back(++externals);
        break;
      case XlsbRecordType::BrtExternSheet: {
        const std::uint32_t count = U32(r.payload, 0);
        for (std::uint32_t i = 0; i < count; ++i) {
          const std::size_t at = 4U + 12U * i;
          const std::uint32_t sup = U32(r.payload, at);
          EXPECT_LT(sup, book_of_sup.size());
          out.emplace_back(sup < book_of_sup.size() ? book_of_sup[sup] : 0xFFFFU,
                           static_cast<std::int32_t>(U32(r.payload, at + 4)),
                           static_cast<std::int32_t>(U32(r.payload, at + 8)));
        }
        break;
      }
      default:
        break;
    }
  }
  return out;
}

/// `rgce` and `rgcb` per row of column A of sheet1.bin.
struct Streams {
  std::vector<std::uint8_t> rgce;
  std::vector<std::uint8_t> rgcb;
};

std::map<std::uint32_t, Streams> ReadStreams(const std::vector<std::uint8_t>& package) {
  std::map<std::uint32_t, Streams> out;
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  auto sheet = zip.read_entry("xl/worksheets/sheet1.bin");
  if (!sheet) {
    ADD_FAILURE() << "sheet1.bin missing";
    return out;
  }
  std::map<std::uint32_t, Streams> arrays;
  std::uint32_t row = 0;
  ByteSpan cursor{sheet.value().data(), sheet.value().size()};
  while (cursor.size > 0U) {
    auto record = read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "unreadable record";
      break;
    }
    const XlsbRecord& r = record.value();
    const ByteSpan p = r.payload;
    if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtRowHdr)) {
      row = U32(p, 0);
    } else if (r.type >= static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaString) &&
               r.type <= static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaError)) {
      std::size_t at = 8;
      if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaString)) {
        at += 4U + 2U * U32(p, at);
      } else if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaNum)) {
        at += 8U;
      } else {
        at += 1U;
      }
      at += 2U;
      const std::uint32_t cce = U32(p, at);
      const std::size_t cb_at = at + 4U + cce;
      const std::uint32_t cb = U32(p, cb_at);
      out[row] = Streams{std::vector<std::uint8_t>(p.data + at + 4, p.data + at + 4 + cce),
                         std::vector<std::uint8_t>(p.data + cb_at + 4, p.data + cb_at + 4 + cb)};
    } else if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtArrFmla)) {
      const std::uint32_t cce = U32(p, 17);
      const std::size_t cb_at = 21U + cce;
      const std::uint32_t cb = U32(p, cb_at);
      arrays[U32(p, 0)] = Streams{std::vector<std::uint8_t>(p.data + 21, p.data + 21 + cce),
                                  std::vector<std::uint8_t>(p.data + cb_at + 4, p.data + cb_at + 4 + cb)};
    }
  }
  for (auto& [key, streams] : arrays) {
    out[key] = std::move(streams);
  }
  return out;
}

/// The leading 3-D token's XTI and the rgce without its `ixti`.
std::pair<Xti, std::vector<std::uint8_t>> Normalised(const std::vector<std::uint8_t>& rgce,
                                                     const std::vector<Xti>& table) {
  if (rgce.size() < 3U || ((rgce[0] & 0x1FU) != 0x1AU && (rgce[0] & 0x1FU) != 0x1BU)) {
    return {Xti{0xFFFFU, 0, 0}, rgce};  // no XTI to rebind
  }
  const std::size_t ixti = static_cast<std::size_t>(rgce[1]) | (static_cast<std::size_t>(rgce[2]) << 8);
  if (ixti >= table.size()) {
    ADD_FAILURE() << "ixti " << ixti << " outside the XTI table";
    return {};
  }
  std::vector<std::uint8_t> rest = rgce;
  rest[1] = rest[2] = 0;
  return {table[ixti], std::move(rest)};
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& package) {
  auto result = read_xlsb(test::span_of(package));
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

Workbook LoadXlsx(const std::vector<std::uint8_t>& package) {
  auto result = read_ooxml(test::span_of(package));
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

void ExpectSameColumnA(const Workbook& got, const Workbook& want, std::uint32_t rows, const char* label) {
  SCOPED_TRACE(label);
  EXPECT_EQ(got.defined_names().size(), want.defined_names().size());
  for (std::size_t i = 0; i < got.defined_names().size() && i < want.defined_names().size(); ++i) {
    EXPECT_EQ(got.defined_names()[i].name, want.defined_names()[i].name);
    EXPECT_EQ(got.defined_names()[i].formula, want.defined_names()[i].formula);
  }
  for (std::uint32_t row = 0; row < rows; ++row) {
    const Cell* g = got.sheet(0).cell_at(row, 0U);
    const Cell* w = want.sheet(0).cell_at(row, 0U);
    ASSERT_NE(g, nullptr) << row;
    ASSERT_NE(w, nullptr) << row;
    EXPECT_EQ(g->formula_text, w->formula_text) << row;
    EXPECT_FALSE(g->formula_text.empty()) << row;
    EXPECT_TRUE(g->cached_value == w->cached_value) << row;
  }
}

TEST(XlsbExtrefDeadSheet, EveryColumnAFormulaReadsAsItsXlsxTwin) {
  for (const FixtureSpec& spec : kFixtures) {
    SCOPED_TRACE(spec.stem);
    const Workbook xlsb = LoadXlsb(test::read_file_bytes(FixturePath(spec, "xlsb")));
    const Workbook xlsx = LoadXlsx(test::read_file_bytes(FixturePath(spec, "xlsx")));
    ExpectSameColumnA(xlsb, xlsx, spec.rows, "xlsb against xlsx");
    ASSERT_EQ(xlsb.defined_names().size(), spec.dead_names);
    for (const DefinedName& name : xlsb.defined_names()) {
      EXPECT_EQ(name.formula, "#REF!");
    }
  }
  const Workbook sheet = LoadXlsb(test::read_file_bytes(FixturePath(kFixtures[0], "xlsb")));
  EXPECT_EQ(sheet.sheet(0).cell_at(0U, 0U)->formula_text, "=#REF!");
  EXPECT_EQ(sheet.sheet(0).cell_at(2U, 0U)->formula_text, "=COUNTA(#REF!)");
  EXPECT_EQ(sheet.sheet(0).cell_at(3U, 0U)->formula_text, "=ISREF(#REF!)");
  const Workbook span = LoadXlsb(test::read_file_bytes(FixturePath(kFixtures[1], "xlsb")));
  EXPECT_EQ(span.sheet(0).cell_at(1U, 0U)->formula_text, "=SUM(#REF!)");
  const Workbook err3d = LoadXlsb(test::read_file_bytes(FixturePath(kFixtures[2], "xlsb")));
  EXPECT_EQ(err3d.sheet(0).cell_at(0U, 0U)->formula_text, "=#REF!");
  EXPECT_EQ(err3d.sheet(0).cell_at(1U, 0U)->formula_text, "=SUM(#REF!)");
  const Workbook middle = LoadXlsb(test::read_file_bytes(FixturePath(kFixtures[4], "xlsb")));
  EXPECT_NE(middle.sheet(0).cell_at(0U, 0U)->formula_text.find("[A.xlsx]S1:S3'!A1)"), std::string::npos);
}

TEST(XlsbExtrefDeadSheet, SaveKeepsTextAndValueAndTheTokensOfEveryLiveReference) {
  for (const FixtureSpec& spec : kFixtures) {
    SCOPED_TRACE(spec.stem);
    const std::vector<std::uint8_t> source = test::read_file_bytes(FixturePath(spec, "xlsb"));
    const Workbook wb = LoadXlsb(source);
    auto saved = write_xlsb(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    ExpectSameColumnA(LoadXlsb(saved.value()), wb, spec.rows, "re-read");

    const std::vector<Xti> source_table = ReadXtiTable(source);
    const std::vector<Xti> saved_table = ReadXtiTable(saved.value());
    const std::map<std::uint32_t, Streams> excel = ReadStreams(source);
    const std::map<std::uint32_t, Streams> ours = ReadStreams(saved.value());
    for (std::uint32_t row = 0; row < spec.rows; ++row) {
      SCOPED_TRACE(row);
      const auto e = excel.find(row);
      const auto o = ours.find(row);
      ASSERT_NE(e, excel.end());
      ASSERT_NE(o, ours.end());
      if (IsDead(spec, row)) {
        // The text model carries no book or address for this form in either format, so it saves as `PtgRefErr`.
        const std::vector<std::uint8_t>& rgce = o->second.rgce;
        ASSERT_GE(rgce.size(), 7U);
        EXPECT_EQ(rgce[0] & 0x1FU, 0x0AU);
        EXPECT_EQ(std::vector<std::uint8_t>(rgce.begin() + 1, rgce.begin() + 7), std::vector<std::uint8_t>(6, 0));
        continue;
      }
      if (std::count(spec.array_rows.begin(), spec.array_rows.end(), row) != 0) {
        continue;
      }
      EXPECT_EQ(Normalised(o->second.rgce, saved_table), Normalised(e->second.rgce, source_table));
      EXPECT_EQ(o->second.rgcb, e->second.rgcb);
    }
  }
}

// A name body reading `#REF!` for a deleted external sheet saves as a zero-location `PtgRefErr`.
TEST(XlsbExtrefDeadSheet, ADeadNameBodySavesAsPtgRefErr) {
  const Workbook wb = LoadXlsb(test::read_file_bytes(FixturePath(kFixtures[3], "xlsb")));
  auto saved = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(test::span_of(saved.value()))));
  auto book = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(book));
  const std::vector<std::uint8_t> body = {7, 0, 0, 0, 0x2A, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0};  // cce, PtgRefErr, cb
  EXPECT_NE(std::search(book.value().begin(), book.value().end(), body.begin(), body.end()), book.value().end());
}

TEST(XlsbExtrefDeadSheet, SaveAsXlsxKeepsTextAndValue) {
  for (const FixtureSpec& spec : kFixtures) {
    SCOPED_TRACE(spec.stem);
    const Workbook wb = LoadXlsb(test::read_file_bytes(FixturePath(spec, "xlsb")));
    auto saved = write_ooxml(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    ExpectSameColumnA(LoadXlsx(saved.value()), wb, spec.rows, "xlsx re-read");
  }
}

// An external sheet name past 31 characters is kept whole in both packages.
TEST(XlsbExtrefDeadSheet, AnExternalSheetNameOver31CharactersSurvivesBothSaves) {
  const std::string name = "Quarterly Revenue Summary Data 26";
  ASSERT_EQ(name.size(), 33U);
  const std::vector<std::uint8_t> source =
      test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_extref_long_sheet_name_B.xlsx");
  const Workbook loaded = LoadXlsx(source);
  const std::string tail = "[A.xlsx]" + name + "'!A1";
  const auto ends_with_name = [&tail](const std::string& text) {
    return text.size() > tail.size() && text.compare(text.size() - tail.size(), tail.size(), tail) == 0;
  };
  const std::string text = loaded.sheet(0).cell_at(8U, 0U)->formula_text;
  EXPECT_TRUE(ends_with_name(text)) << text;

  auto xlsx = write_ooxml(loaded);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  std::string sheet_xml;
  std::string link_xml;
  ASSERT_TRUE(test::extract_part(test::span_of(xlsx.value()), "xl/worksheets/sheet1.xml", &sheet_xml));
  ASSERT_TRUE(test::extract_part(test::span_of(xlsx.value()), "xl/externalLinks/externalLink1.xml", &link_xml));
  EXPECT_NE(sheet_xml.find("<f>&apos;[1]" + name + "&apos;!A1</f>"), std::string::npos);
  EXPECT_NE(link_xml.find("<sheetName val=\"" + name + "\"/>"), std::string::npos);

  auto xlsb = write_xlsb(loaded);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(test::span_of(xlsb.value()))));
  auto link = zip.read_entry("xl/externalLinks/externalLink1.bin");
  ASSERT_TRUE(static_cast<bool>(link));
  std::vector<std::uint8_t> sheet_name = {33, 0, 0, 0};  // BrtSupTabs entry: cch, then UTF-16LE
  for (const char c : name) {
    sheet_name.push_back(static_cast<std::uint8_t>(c));
    sheet_name.push_back(0);
  }
  EXPECT_NE(std::search(link.value().begin(), link.value().end(), sheet_name.begin(), sheet_name.end()),
            link.value().end());

  const Workbook back = LoadXlsb(xlsb.value());
  EXPECT_EQ(back.sheet(0).cell_at(8U, 0U)->formula_text, text);
  ExpectSameColumnA(back, loaded, 9, "xlsb re-read");
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
