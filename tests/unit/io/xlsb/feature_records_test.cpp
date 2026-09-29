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
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
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
/// only an .xlsx source can carry. The x14 overlay of a linked data bar is
/// not decoded from .xlsb, so its lengths -- which the overlay overrides --
/// are left out for linked rules.
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
        if (r.id.empty()) {
          os << " len=" << static_cast<int>(r.data_bar->min_length_pct) << "-"
             << static_cast<int>(r.data_bar->max_length_pct);
        }
        os << " show=" << r.data_bar->show_value;
      }
      if (r.icon_set) {
        os << " icons=" << static_cast<int>(r.icon_set->name) << " rev=" << r.icon_set->reverse
           << " show=" << r.icon_set->show_value << " pct=" << r.icon_set->percent;
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
    const bool feature = (t >= 461 && t <= 471) || t == 564 || t == 1146 || t == 573 || t == 574 || t == 64 ||
                         t == 535 || t == 678 || t == 534 || t == 677 || ((t == 35 || t == 36) && !in_dvals);
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
                                           "dv_all", "prot", "prot2", "excelprot", "x14"));

// Left out of the byte comparison, all for the Ptg encoder's canonical
// form or the model's shape rather than the record layout: `flags` and
// `text_rules` (Excel adds PtgParen, a volatile PtgAttrSemi and IF's
// PtgAttr jumps, which the encoder never emits), `rel` (the same, plus
// its own function-token classes) and `iconbits` (a `3Flags` set whose
// `num 0` floor threshold the model drops, as the OOXML path does).
INSTANTIATE_TEST_SUITE_P(Excel, XlsbFeatureWriterBytes,
                         ::testing::Values("base", "cellis_ops", "cfvo", "dv_all", "prot", "prot2", "excelprot",
                                           "x14"));

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

TEST(XlsbFeatureRecords, XlsxSourcedFeaturesAreWrittenToXlsb) {
  const Workbook from_xlsx = ReadXlsx("base");
  auto written = io::xlsb::write_xlsb(from_xlsx);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  const Workbook back = ReadXlsbBytes(written.value());
  EXPECT_EQ(DescribeCf(back.sheet(0)), DescribeCf(from_xlsx.sheet(0)));
  EXPECT_EQ(DescribeDv(back.sheet(0)), DescribeDv(from_xlsx.sheet(0)));
}

}  // namespace
}  // namespace formulon
