//
// Conditional formatting, data validation and protection on the MS-XLSB
// path, against real Mac Excel 365 pairs
// (`tests/fixtures/excel/xlsb_feature_*.{xlsb,xlsx}`).
//
// Each pair was authored as .xlsx with openpyxl (the protected-with-hash
// pair by Excel itself), opened in Excel and saved by Excel as both .xlsb
// and .xlsx. Excel's .xlsx is the oracle: the binary reader must decode
// to the model the OOXML reader builds from it. The writer is then held to
// Excel's own bytes for the records it regenerates.
//
//   base       CF (colour scale, data bar, icon set, cellIs, expression on a
//              two-range sqref) and DV (whole with messages, inline list,
//              range list, custom) on one sheet
//   cellis_ops every cellIs operator
//   text_rules text / blanks / errors / duplicate / unique / timePeriod /
//              top10 / aboveAverage rules
//   flags      top10 and aboveAverage flag variants, every timePeriod
//   cfvo       every legacy threshold type, icon-set options, data bar
//              lengths and showValue
//   iconbits   icon-set gte bits; one rule Excel keeps only as x14
//   rel        PtgRefN / PtgAreaN bases, a defined name, literals
//   dv_all     every DV type x operator, error style, message flags
//   prot       legacy sheet password, flag sets, workbook lockStructure
//   prot2      the remaining protection flags, one per sheet
//   excelprot  SHA-512 sheet and workbook passwords set by Excel
//   x14        data bars linked to x14 counterparts
//   x14bars    x14 data-bar fields varied one per rule: border, gradient,
//              axis position and colour, negative fill and border, lengths
//   x14dir     each x14 data-bar direction: leftToRight, context, rightToLeft
//   dxf        one CF dxf per property: font toggles, underline kinds, colour
//              kinds, fill patterns, border sides and diagonals
//   dxfnum     CF dxfs with number formats, alone and with a font
//   dxfhand    CF dxfs with a font name, a size, alignment and protection;
//              Excel keeps only the name

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <optional>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/styles_writer.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
  }
  std::fclose(file);
  return out;
}

std::string FixturePath(const std::string& name, const char* ext) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_feature_" + name + "." + ext;
}

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

Workbook ReadXlsbBytes(const std::vector<std::uint8_t>& bytes) {
  auto result = io::xlsb::read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

Workbook ReadXlsx(const std::string& name) {
  const std::vector<std::uint8_t> bytes = ReadFileBytes(FixturePath(name, "xlsx"));
  auto result = io::read_ooxml(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

/// A formula in canonical form. The Ptg codec does not keep redundant
/// parentheses (PtgParen is transparent, as for cell formulas), so text
/// is compared after one parse / format pass.
std::string Canonical(const std::string& text) {
  if (text.empty()) {
    return text;
  }
  Arena arena;
  parser::Parser parser(text, arena);
  const parser::AstNode* root = parser.parse();
  return root != nullptr && parser.errors().empty() ? parser::format_formula(*root) : "unparsed:" + text;
}

std::string Canonical(const std::optional<std::string>& text) {
  return text ? Canonical(*text) : "-";
}

void DescribeCfvo(std::ostream& os, const cf::CfValueObject& v) {
  os << "{" << static_cast<int>(v.type) << ":" << v.value << (v.gte ? "" : " gt") << "}";
}

void DescribeColor(std::ostream& os, cf::Color c) {
  os << "#" << static_cast<int>(c.a) << "." << static_cast<int>(c.r) << "." << static_cast<int>(c.g) << "."
     << static_cast<int>(c.b);
}

/// Every modelled CF field except the verbatim `ext_lst_raw` XML, which
/// only an .xlsx source can carry.
std::string DescribeCf(const Sheet& sheet) {
  std::ostringstream os;
  for (const cf::ConditionalFormat& f : sheet.conditional_formats()) {
    os << "[";
    for (const cf::CFCellRange& r : f.sqref) {
      os << r.first.row << "," << r.first.col << ":" << r.last.row << "," << r.last.col << " ";
    }
    os << "pivot=" << f.pivot_scope << "\n";
    for (const cf::CFRule& r : f.rules) {
      os << "  type=" << static_cast<int>(r.type) << " pri=" << r.priority << " stop=" << r.stop_if_true
         << " dxf=" << (r.dxf_id ? static_cast<long>(*r.dxf_id) : -1L) << " id=" << r.id
         << " f1=" << Canonical(r.formula1) << " f2=" << Canonical(r.formula2)
         << " op=" << (r.op ? static_cast<int>(*r.op) : -1) << " rank=" << r.rank.value_or(-1) << " pct=" << r.percent
         << " bottom=" << r.bottom << " above=" << r.above_average << " eq=" << r.equal_average
         << " sd=" << r.std_dev.value_or(-1) << " text=" << r.text.value_or("-")
         << " tp=" << (r.time_period ? static_cast<int>(*r.time_period) : -1);
      if (r.color_scale) {
        os << " scale";
        for (const auto& v : r.color_scale->thresholds) {
          DescribeCfvo(os, v);
        }
        for (cf::Color c : r.color_scale->colors) {
          DescribeColor(os, c);
        }
      }
      if (r.data_bar) {
        os << " bar";
        DescribeCfvo(os, r.data_bar->min);
        DescribeCfvo(os, r.data_bar->max);
        DescribeColor(os, r.data_bar->fill);
        const cf::DataBarSpec& b = *r.data_bar;
        os << " len=" << static_cast<int>(b.min_length_pct) << "-" << static_cast<int>(b.max_length_pct)
           << " show=" << b.show_value << " grad=" << b.gradient << " axis=" << static_cast<int>(b.axis_position)
           << " dir=" << static_cast<int>(b.direction);
        os << " neg";
        DescribeColor(os, b.negative_fill);
        os << " axiscol";
        DescribeColor(os, b.axis_color);
        if (b.border) {
          os << " border";
          DescribeColor(os, *b.border);
        }
        if (b.negative_border) {
          os << " negborder";
          DescribeColor(os, *b.negative_border);
        }
      }
      if (r.icon_set) {
        os << " icons=" << static_cast<int>(r.icon_set->name) << " rev=" << r.icon_set->reverse
           << " show=" << r.icon_set->show_value << " pct=" << r.icon_set->percent << " floor";
        DescribeCfvo(os, r.icon_set->floor);
        for (const auto& v : r.icon_set->thresholds) {
          DescribeCfvo(os, v);
        }
      }
      os << "\n";
    }
    os << "]\n";
  }
  return os.str();
}

std::string DescribeDv(const Sheet& sheet) {
  std::ostringstream os;
  for (const DataValidation& v : sheet.validations()) {
    for (const MergeRange& r : v.ranges) {
      os << r.first_row << "," << r.first_col << ":" << r.last_row << "," << r.last_col << " ";
    }
    os << "t=" << static_cast<int>(v.type) << " op=" << static_cast<int>(v.op)
       << " err=" << static_cast<int>(v.error_style) << " blank=" << v.allow_blank << " in=" << v.show_input_message
       << " er=" << v.show_error_message << " dd=" << v.show_dropdown << " f1=" << Canonical(v.formula1)
       << " f2=" << Canonical(v.formula2) << " et=" << v.error_title << " em=" << v.error_message
       << " pt=" << v.prompt_title << " pm=" << v.prompt_message << "\n";
  }
  return os.str();
}

std::string DescribeProtection(const Sheet& sheet) {
  const SheetProtection& p = sheet.protection();
  std::ostringstream os;
  os << "on=" << p.enabled;
  if (!p.enabled) {
    return os.str();
  }
  os << " alg=" << p.algorithm_name << " hash=" << p.hash_value << " salt=" << p.salt_value << " spin=" << p.spin_count
     << " pwd=" << p.legacy_password << " flags=" << p.sheet << p.objects << p.scenarios << p.format_cells
     << p.format_columns << p.format_rows << p.insert_columns << p.insert_rows << p.insert_hyperlinks
     << p.delete_columns << p.delete_rows << p.select_locked_cells << p.select_unlocked_cells << p.sort << p.auto_filter
     << p.pivot_tables;
  return os.str();
}

std::string DescribeFeatures(const Workbook& wb) {
  std::string out = "book=" + wb.workbook_protection_xml() + "\n";
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    const Sheet& s = wb.sheet(i);
    out += "sheet " + std::to_string(i) + " " + DescribeProtection(s) + "\n" + DescribeCf(s) + DescribeDv(s);
  }
  return out;
}

/// Tail records of `part` of the kinds this change regenerates, one
/// `type: hex payload` line each. The alternate-content uid wrappers
/// around each BrtDVal are revision ids the model drops.
std::vector<std::string> FeatureRecords(const std::vector<std::uint8_t>& xlsb, const std::string& part) {
  std::vector<std::string> out;
  io::ZipReader zip;
  if (!zip.open(SpanOf(xlsb))) {
    ADD_FAILURE() << "zip open failed";
    return out;
  }
  auto body = zip.read_entry(part);
  if (!body) {
    ADD_FAILURE() << "missing part " << part;
    return out;
  }
  bool in_dvals = false;
  bool in_tail = false;
  io::ByteSpan cursor = SpanOf(body.value());
  while (cursor.size != 0U) {
    auto rec = io::xlsb::read_record(cursor);
    if (!rec) {
      ADD_FAILURE() << "record walk failed";
      break;
    }
    const std::uint16_t t = rec.value().type;
    in_dvals = t == 573 ? true : (t == 574 ? false : in_dvals);
    if (t == 146) {
      in_tail = true;
    }
    if (!in_tail) {
      continue;
    }
    const bool x14_cf = t == 1135 || t == 1136 || (t >= 1046 && t <= 1051) || t == 1055 || t == 1156;
    const bool feature = (t >= 461 && t <= 471) || t == 564 || t == 1146 || t == 573 || t == 574 || t == 64 ||
                         t == 535 || t == 678 || t == 534 || t == 677 || x14_cf || ((t == 35 || t == 36) && !in_dvals);
    if (feature) {
      std::string line = std::to_string(t) + ":";
      for (std::size_t i = 0; i < rec.value().payload.size; ++i) {
        char hex[4];
        std::snprintf(hex, sizeof(hex), " %02x", rec.value().payload.data[i]);
        line += hex;
      }
      out.push_back(std::move(line));
    }
  }
  return out;
}

/// True when `a` and `b` differ at most in formula size fields and in the
/// class bits of reference Ptg opcodes (Ref / Area / Name and their N and
/// 3-D forms). The Ptg encoder assigns operand classes its own way, and
/// Excel re-derives them on load: its re-save of such a file matches its
/// own original.
bool SameUpToPtgClass(const std::string& a, const std::string& b) {
  const std::size_t colon = a.find(':');
  if (a.size() != b.size() || colon == std::string::npos || a.compare(0, colon, b, 0, colon) != 0) {
    return false;
  }
  // The formula size fields (BrtBeginCFRule bytes 30-41, BrtCFVO 20-23),
  // which the writer fills with `cce` (see `EncodedFeatureFormula`).
  const std::string type = a.substr(0, colon);
  const std::size_t skip_from = type == "463" ? 30U : type == "471" ? 20U : 0U;
  const std::size_t skip_to = type == "463" ? 42U : type == "471" ? 24U : 0U;
  for (std::size_t i = colon + 2U; i + 1U < a.size(); i += 3U) {
    const std::size_t byte = (i - colon - 2U) / 3U;
    if (a.compare(i, 2, b, i, 2) == 0 || (byte >= skip_from && byte < skip_to)) {
      continue;
    }
    const auto x = static_cast<unsigned>(std::stoul(a.substr(i, 2), nullptr, 16));
    const auto y = static_cast<unsigned>(std::stoul(b.substr(i, 2), nullptr, 16));
    const unsigned base = (x & 0x1FU) | 0x20U;
    const bool reference = (base >= 0x23U && base <= 0x2DU) || (base >= 0x3AU && base <= 0x3DU);
    if (((x ^ y) & ~0x60U) != 0U || x < 0x20U || x >= 0x80U || y < 0x20U || !reference) {
      return false;
    }
  }
  return true;
}

std::string Joined(const std::vector<std::string>& lines) {
  std::string out;
  for (const std::string& line : lines) {
    out += line + "\n";
  }
  return out;
}

class XlsbFeatureFixture : public ::testing::TestWithParam<const char*> {};

TEST_P(XlsbFeatureFixture, DecodesToTheModelExcelsXlsxYields) {
  const std::string name = GetParam();
  const Workbook from_xlsb = ReadXlsbBytes(ReadFileBytes(FixturePath(name, "xlsb")));
  const Workbook from_xlsx = ReadXlsx(name);
  EXPECT_EQ(DescribeFeatures(from_xlsb), DescribeFeatures(from_xlsx));
}

TEST_P(XlsbFeatureFixture, RoundTripsThroughTheModel) {
  const std::string name = GetParam();
  const Workbook wb = ReadXlsbBytes(ReadFileBytes(FixturePath(name, "xlsb")));
  auto written = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  EXPECT_EQ(DescribeFeatures(ReadXlsbBytes(written.value())), DescribeFeatures(wb));
}

/// The regenerated tail records match Excel's own bytes, record for record.
/// Excel writes a BrtSheetProtection for every sheet, protected or not;
/// the writer emits one only for a protected sheet, so unprotected sheets'
/// all-default records are left out of the comparison.
class XlsbFeatureWriterBytes : public ::testing::TestWithParam<const char*> {};

TEST_P(XlsbFeatureWriterBytes, MatchExcelUpToPtgClass) {
  const std::string name = GetParam();
  const std::vector<std::uint8_t> source = ReadFileBytes(FixturePath(name, "xlsb"));
  const Workbook wb = ReadXlsbBytes(source);
  auto written = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    const std::string part = "xl/worksheets/sheet" + std::to_string(i + 1) + ".bin";
    auto expected = FeatureRecords(source, part);
    if (!wb.sheet(i).protection().enabled) {
      expected.erase(std::remove_if(expected.begin(), expected.end(),
                                    [](const std::string& r) { return r.rfind("535:", 0) == 0; }),
                     expected.end());
    }
    const std::vector<std::string> actual = FeatureRecords(written.value(), part);
    bool same = actual.size() == expected.size();
    for (std::size_t r = 0; same && r < actual.size(); ++r) {
      same = SameUpToPtgClass(actual[r], expected[r]);
    }
    if (!same) {
      EXPECT_EQ(Joined(actual), Joined(expected)) << part;
    }
  }
}

INSTANTIATE_TEST_SUITE_P(Excel, XlsbFeatureFixture,
                         ::testing::Values("base", "cellis_ops", "text_rules", "flags", "cfvo", "iconbits", "rel",
                                           "dv_all", "prot", "prot2", "excelprot", "x14", "x14bars", "x14dir"));

// Left out: `text_rules`, `flags` and `rel`, whose formulas differ from
// Excel's only in the Ptg codec's canonical form, not in record layout:
// redundant parentheses the decoder does not keep, the PtgAttrSemi operand
// (`00 00` where Excel writes `fe ff`), and IF's PtgAttr jumps.
INSTANTIATE_TEST_SUITE_P(Excel, XlsbFeatureWriterBytes,
                         ::testing::Values("base", "cellis_ops", "cfvo", "iconbits", "dv_all", "prot", "prot2",
                                           "excelprot", "x14", "x14bars", "x14dir"));

/// Payload of the first `type` record in `part`, hex-encoded.
std::string RecordPayload(const std::vector<std::uint8_t>& xlsb, const std::string& part, std::uint16_t type) {
  io::ZipReader zip;
  if (!zip.open(SpanOf(xlsb))) {
    return "no zip";
  }
  auto body = zip.read_entry(part);
  if (!body) {
    return "no part";
  }
  io::ByteSpan cursor = SpanOf(body.value());
  while (cursor.size != 0U) {
    auto rec = io::xlsb::read_record(cursor);
    if (!rec) {
      break;
    }
    if (rec.value().type == type) {
      std::string out;
      for (std::size_t i = 0; i < rec.value().payload.size; ++i) {
        char hex[4];
        std::snprintf(hex, sizeof(hex), "%02x ", rec.value().payload.data[i]);
        out += hex;
      }
      return out;
    }
  }
  return "absent";
}

// `iconbits` and `rel` each keep an x14 block Excel writes only in its
// extension form, whose formula names sheet S2 through a BrtExternSheet
// index. The written table has to keep that index naming S2, or Excel
// refuses the file.
TEST(XlsbFeatureRecords, RetainedX14SheetReferencesKeepTheirIndices) {
  for (const char* name : {"iconbits", "rel"}) {
    const std::vector<std::uint8_t> source = ReadFileBytes(FixturePath(name, "xlsb"));
    auto written = io::xlsb::write_xlsb(ReadXlsbBytes(source));
    ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
    EXPECT_EQ(RecordPayload(written.value(), "xl/workbook.bin", 362), RecordPayload(source, "xl/workbook.bin", 362))
        << name;
  }
}

TEST(XlsbFeatureRecords, RetainedX14SheetReferenceToARenamedSheetFailsClosed) {
  Workbook wb = ReadXlsbBytes(ReadFileBytes(FixturePath("iconbits", "xlsb")));
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(1U, "Renamed")));
  auto written = io::xlsb::write_xlsb(wb);
  ASSERT_FALSE(static_cast<bool>(written));
  EXPECT_EQ(written.error().code, FormulonErrorCode::kIoXlsbRetainedPartStale);
}

std::size_t CountRecords(const std::vector<std::uint8_t>& xlsb, const std::string& part, std::uint16_t type) {
  io::ZipReader zip;
  if (!zip.open(SpanOf(xlsb))) {
    return 0;
  }
  auto body = zip.read_entry(part);
  std::size_t count = 0;
  io::ByteSpan cursor = body ? SpanOf(body.value()) : io::ByteSpan{};
  while (cursor.size != 0U) {
    auto rec = io::xlsb::read_record(cursor);
    if (!rec) {
      break;
    }
    count += rec.value().type == type ? 1U : 0U;
  }
  return count;
}

// The model wins over a retained x14 data bar: edited settings are
// written, a deleted rule's x14 half goes with it.
TEST(XlsbFeatureRecords, EditedX14DataBarsReachTheSavedFile) {
  Workbook wb = ReadXlsbBytes(ReadFileBytes(FixturePath("x14bars", "xlsb")));
  std::vector<cf::ConditionalFormat>& formats = wb.sheet(0).mutable_conditional_formats();
  ASSERT_EQ(formats.size(), 5U);
  cf::DataBarSpec& bar = *formats[0].rules[0].data_bar;
  bar.negative_fill = cf::Color{0x12, 0x34, 0x56, 255};
  bar.axis_position = cf::DataBarAxisPosition::Middle;
  bar.gradient = false;
  bar.max_length_pct = 80;
  formats.erase(formats.begin() + 1);
  auto written = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  const Workbook back = ReadXlsbBytes(written.value());
  EXPECT_EQ(DescribeFeatures(back), DescribeFeatures(wb));
  EXPECT_EQ(CountRecords(written.value(), "xl/worksheets/sheet1.bin", 1048), 4U);
}

// A model data bar with x14-only settings and no retained x14 record (an
// .xlsx source here) gets one written.
TEST(XlsbFeatureRecords, X14DataBarsFromAnXlsxSourceAreWritten) {
  const Workbook from_xlsx = ReadXlsx("x14bars");
  auto written = io::xlsb::write_xlsb(from_xlsx);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  EXPECT_EQ(DescribeCf(ReadXlsbBytes(written.value()).sheet(0)), DescribeCf(from_xlsx.sheet(0)));
}

TEST(XlsbFeatureRecords, EditedX14DataBarsReachASavedXlsx) {
  Workbook wb = ReadXlsx("x14bars");
  cf::DataBarSpec& bar = *wb.sheet(0).mutable_conditional_formats()[0].rules[0].data_bar;
  bar.negative_fill = cf::Color{0x12, 0x34, 0x56, 255};
  bar.axis_position = cf::DataBarAxisPosition::None;
  auto saved = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto back = io::read_ooxml(SpanOf(saved.value()));
  ASSERT_TRUE(static_cast<bool>(back)) << back.error().message;
  EXPECT_EQ(DescribeCf(back.value().workbook.sheet(0)), DescribeCf(wb.sheet(0)));
}

/// Every BrtDXF payload of a styles part, hex-encoded. Excel writes the
/// resolved RGB beside a theme, indexed or automatic colour; the model does
/// not carry it, so those four bytes are masked.
std::vector<std::string> DxfRecords(io::ByteSpan styles) {
  std::vector<std::string> out;
  while (styles.size != 0U) {
    auto rec = io::xlsb::read_record(styles);
    if (!rec) {
      ADD_FAILURE() << "record walk failed";
      break;
    }
    if (rec.value().type == 507) {
      std::vector<std::uint8_t> p(rec.value().payload.data, rec.value().payload.data + rec.value().payload.size);
      static constexpr std::uint16_t kColorProps[] = {1, 2, 5, 6, 7, 8, 9, 10};
      for (std::size_t at = 6; at + 4U <= p.size();) {
        const std::uint16_t type = static_cast<std::uint16_t>(p[at] | (p[at + 1] << 8U));
        const std::uint16_t size = static_cast<std::uint16_t>(p[at + 2] | (p[at + 3] << 8U));
        const bool color = std::find(std::begin(kColorProps), std::end(kColorProps), type) != std::end(kColorProps);
        if (color && size >= 12U && (p[at + 4] >> 1U) != 2U) {
          std::fill(p.begin() + static_cast<std::ptrdiff_t>(at + 8), p.begin() + static_cast<std::ptrdiff_t>(at + 12),
                    std::uint8_t{0});
        }
        at += size == 0U ? p.size() : size;
      }
      std::string line;
      for (std::uint8_t byte : p) {
        char hex[4];
        std::snprintf(hex, sizeof(hex), "%02x ", byte);
        line += hex;
      }
      out.push_back(std::move(line));
    }
  }
  return out;
}

// The dxfs Excel wrote into its .xlsx, written to XLSB, match the BrtDXF
// records Excel wrote into the .xlsb of the same book.
TEST(XlsbFeatureRecords, DxfRecordsMatchExcel) {
  for (const char* name : {"dxf", "dxfnum", "dxfhand"}) {
    const Workbook from_xlsx = ReadXlsx(name);
    const std::vector<std::uint8_t> written = io::xlsb::write_styles_bin(from_xlsx.styles());
    io::ZipReader zip;
    const std::vector<std::uint8_t> excel = ReadFileBytes(FixturePath(name, "xlsb"));
    ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(excel))));
    auto styles = zip.read_entry("xl/styles.bin");
    ASSERT_TRUE(static_cast<bool>(styles));
    EXPECT_EQ(Joined(DxfRecords(SpanOf(written))), Joined(DxfRecords(SpanOf(styles.value())))) << name;
  }
}

// CF and DV formulas are held in formula-bar spelling, like cell formulas:
// storage prefixes stripped from known names, SINGLE / ANCHORARRAY shown as
// `@` / `#`, an unknown `_xlfn.` name left as is. The .xlsx writer spells
// them back the way Excel stores them.
constexpr const char* kCanonicalFeatureFormula = "AND(XLOOKUP(A1,B1:B3,C1:C3)>0,@A1>0,SUM(A1#)>0,_xlfn.FOOBAR(1))";
constexpr const char* kStoredFeatureFormula =
    "AND(_xlfn.XLOOKUP(A1,B1:B3,C1:C3)&gt;0,_xlfn.SINGLE(A1)&gt;0,SUM(_xlfn.ANCHORARRAY(A1))&gt;0,_xlfn.FOOBAR(1))";

Workbook FeatureFormulaWorkbook() {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("S1"));
  cf::ConditionalFormat format;
  format.sqref.push_back(cf::CFCellRange{{0, 3}, {4, 3}});
  cf::CFRule expression;
  expression.type = cf::RuleType::Expression;
  expression.formula1 = kCanonicalFeatureFormula;
  format.rules.push_back(expression);
  cf::CFRule scale;
  scale.type = cf::RuleType::ColorScale;
  scale.priority = 2;
  cf::ColorScaleSpec spec;
  spec.thresholds = {cf::CfValueObject{cf::CfvoType::Min, "", true},
                     cf::CfValueObject{cf::CfvoType::Formula, "@A1", true}};
  spec.colors = {cf::Color{1, 2, 3, 255}, cf::Color{4, 5, 6, 255}};
  scale.color_scale = spec;
  format.rules.push_back(scale);
  sheet.mutable_conditional_formats().push_back(format);
  DataValidation dv;
  dv.ranges.push_back(MergeRange{0, 4, 4, 4});
  dv.type = 7;  // custom
  dv.formula1 = kCanonicalFeatureFormula;
  sheet.mutable_validations().push_back(dv);
  return wb;
}

void ExpectCanonicalFeatureFormulas(const Workbook& wb, const char* where) {
  const Sheet& sheet = wb.sheet(0);
  ASSERT_EQ(sheet.conditional_formats().size(), 1U) << where;
  const cf::ConditionalFormat& format = sheet.conditional_formats()[0];
  ASSERT_EQ(format.rules.size(), 2U) << where;
  EXPECT_EQ(format.rules[0].formula1.value_or(""), kCanonicalFeatureFormula) << where;
  ASSERT_TRUE(format.rules[1].color_scale.has_value()) << where;
  EXPECT_EQ(format.rules[1].color_scale->thresholds[1].value, "@A1") << where;
  ASSERT_EQ(sheet.validations().size(), 1U) << where;
  EXPECT_EQ(sheet.validations()[0].formula1, kCanonicalFeatureFormula) << where;
}

void ExpectStoredFeatureFormulas(const std::vector<std::uint8_t>& xlsx, const char* where) {
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(xlsx)))) << where;
  auto part = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(part)) << where;
  const std::string xml(part.value().begin(), part.value().end());
  EXPECT_NE(xml.find(std::string("<formula>") + kStoredFeatureFormula + "</formula>"), std::string::npos)
      << where << "\n"
      << xml;
  EXPECT_NE(xml.find(std::string("<formula1>") + kStoredFeatureFormula + "</formula1>"), std::string::npos)
      << where << "\n"
      << xml;
  EXPECT_NE(xml.find("<cfvo type=\"formula\" val=\"_xlfn.SINGLE(A1)\"/>"), std::string::npos) << where << "\n" << xml;
}

TEST(XlsbFeatureRecords, CfAndDvFormulasAreCanonicalInTheModel) {
  auto first = io::write_ooxml(FeatureFormulaWorkbook());
  ASSERT_TRUE(static_cast<bool>(first)) << first.error().message;
  ExpectStoredFeatureFormulas(first.value(), "xlsx");
  auto loaded = io::read_ooxml(SpanOf(first.value()));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  ExpectCanonicalFeatureFormulas(loaded.value().workbook, "xlsx -> model");

  auto xlsb = io::xlsb::write_xlsb(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  const Workbook from_xlsb = ReadXlsbBytes(xlsb.value());
  ExpectCanonicalFeatureFormulas(from_xlsb, "xlsx -> xlsb -> model");
  auto again = io::write_ooxml(from_xlsb);
  ASSERT_TRUE(static_cast<bool>(again)) << again.error().message;
  ExpectStoredFeatureFormulas(again.value(), "xlsx -> xlsb -> xlsx");
}

TEST(XlsbFeatureRecords, WorkbookProtectionMatchesExcelsXlsxElement) {
  const Workbook legacy = ReadXlsbBytes(ReadFileBytes(FixturePath("prot", "xlsb")));
  EXPECT_EQ(legacy.workbook_protection_xml(), "<workbookProtection lockStructure=\"1\"/>");
  const Workbook hashed = ReadXlsbBytes(ReadFileBytes(FixturePath("excelprot", "xlsb")));
  EXPECT_EQ(hashed.workbook_protection_xml(),
            "<workbookProtection workbookAlgorithmName=\"SHA-512\" "
            "workbookHashValue=\"YhLLtvALkasOXe1ZYvzxNtMgaAj2ywh3OdVST63BDWPewtyR2NZPdIBHYYguN/5vxAfbiBRXywV69D95/"
            "ehrpQ==\" workbookSaltValue=\"1Ha2LwKiFItTdyxsunmdgQ==\" workbookSpinCount=\"100000\" "
            "lockStructure=\"1\"/>");
}

TEST(XlsbFeatureRecords, ModelEditsReachTheSavedFile) {
  Workbook wb = ReadXlsbBytes(ReadFileBytes(FixturePath("base", "xlsb")));
  Sheet& sheet = wb.sheet(0);
  ASSERT_FALSE(sheet.conditional_formats().empty());
  ASSERT_FALSE(sheet.validations().empty());
  sheet.mutable_conditional_formats().pop_back();
  sheet.mutable_conditional_formats().front().rules.front().priority = 42;
  sheet.mutable_validations().front().prompt_message = "edited";
  sheet.mutable_protection().enabled = true;
  sheet.mutable_protection().sheet = true;
  sheet.mutable_protection().legacy_password = "CC1A";
  auto written = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  const Workbook back = ReadXlsbBytes(written.value());
  EXPECT_EQ(DescribeFeatures(back), DescribeFeatures(wb));
  EXPECT_EQ(back.sheet(0).conditional_formats().front().rules.front().priority, 42);
  EXPECT_EQ(back.sheet(0).protection().legacy_password, "CC1A");
}

// Every CF block of an .xlsx source reaches the .xlsb, the ones whose rules
// reference a dxf included: the generated styles part carries the dxfs.
TEST(XlsbFeatureRecords, XlsxSourcedFeaturesAreWrittenToXlsb) {
  const Workbook from_xlsx = ReadXlsx("base");
  auto written = io::xlsb::write_xlsb_with_result(from_xlsx);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  const Workbook back = ReadXlsbBytes(written.value().bytes);
  EXPECT_EQ(DescribeCf(back.sheet(0)), DescribeCf(from_xlsx.sheet(0)));
  EXPECT_EQ(DescribeDv(back.sheet(0)), DescribeDv(from_xlsx.sheet(0)));
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(written.value().bytes))));
  auto styles = zip.read_entry("xl/styles.bin");
  ASSERT_TRUE(static_cast<bool>(styles));
  EXPECT_EQ(DxfRecords(SpanOf(styles.value())).size(), from_xlsx.styles().dxfs.size());
}

}  // namespace
}  // namespace formulon
