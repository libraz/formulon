// The apply flags a save writes for a cell xf, against Excel 365 (Mac, ja-JP UI)
// re-saving the same files as .xlsx and .xlsb. Every expectation is read from
// Excel's own `.out` files next to the authored input.

#include <cstdint>
#include <memory>
#include <set>
#include <string>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "pugixml.hpp"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io {
namespace {

using Bytes = std::vector<std::uint8_t>;

const std::vector<std::string>& CaseIds() {
  static const std::vector<std::string> ids = {
      "nf_noflag",      "nf_flag0",        "nf_flag1",          "font_noflag",      "font_flag0",
      "fill_noflag",    "fill_flag0",      "border_noflag",     "border_flag0",     "align_noflag",
      "align_flag0",    "prot_noflag",     "prot_flag0",        "same_all1",        "same_nf1",
      "S1_equal_flag1", "S1_equal_noflag", "S1_differs_noflag", "S1_differs_flag0", "all_groups_noflag",
  };
  return ids;
}

std::string CaseName(const ::testing::TestParamInfo<std::string>& info) {
  std::string out;
  for (const char ch : info.param) {
    out.push_back((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9') ? ch : '_');
  }
  return out;
}

std::string FixturePath(const std::string& id, const std::string& suffix) {
  return std::string(FORMULON_FIXTURES_DIR) + "/cellxfs_apply_flags/af_" + id + suffix;
}

std::string PartText(const Bytes& pkg, const std::string& name) {
  std::string text;
  (void)test::extract_part(test::span_of(pkg), name, &text);
  return text;
}

std::unique_ptr<Workbook> LoadAuthored(const std::string& id) {
  const Bytes bytes = test::read_file_bytes(FixturePath(id, ".xlsx"));
  auto read = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  if (!read) {
    ADD_FAILURE() << "read_ooxml " << id << ": " << read.error().message;
    return nullptr;
  }
  auto wb = std::make_unique<Workbook>(std::move(read.value().workbook));
  wb->set_excel_profile(mac_365_ja_jp_profile());
  return wb;
}

/// "NFABPL" style string: the six apply flags in the order numFmt, font, alignment, border, fill, protection,
/// each its letter when set and '-' when not.
std::string FlagString(bool nf, bool font, bool align, bool border, bool fill, bool prot) {
  return std::string{nf ? 'N' : '-',     font ? 'F' : '-', align ? 'A' : '-',
                     border ? 'B' : '-', fill ? 'L' : '-', prot ? 'P' : '-'};
}

std::string XlsxFlags(const pugi::xml_node xf) {
  const auto set = [&xf](const char* name) { return xf.attribute(name).as_bool(); };
  return FlagString(set("applyNumberFormat"), set("applyFont"), set("applyAlignment"), set("applyBorder"),
                    set("applyFill"), set("applyProtection"));
}

struct BookFlags {
  std::string a1;
  std::set<std::string> style_xfs;
};

BookFlags ReadXlsx(const Bytes& pkg) {
  BookFlags out;
  pugi::xml_document styles;
  const std::string styles_xml = PartText(pkg, "xl/styles.xml");
  EXPECT_TRUE(styles.load_buffer(styles_xml.data(), styles_xml.size()));
  pugi::xml_document sheet;
  const std::string sheet_xml = PartText(pkg, "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(sheet.load_buffer(sheet_xml.data(), sheet_xml.size()));
  const pugi::xml_node root = styles.child("styleSheet");
  for (const pugi::xml_node xf : root.child("cellStyleXfs").children("xf")) {
    out.style_xfs.insert(XlsxFlags(xf));
  }
  for (const pugi::xml_node c : sheet.child("worksheet").child("sheetData").child("row").children("c")) {
    if (std::string(c.attribute("r").value()) == "A1") {
      unsigned index = c.attribute("s").as_uint();
      for (const pugi::xml_node xf : root.child("cellXfs").children("xf")) {
        if (index-- == 0U) {
          out.a1 = XlsxFlags(xf);
          break;
        }
      }
    }
  }
  return out;
}

constexpr std::uint16_t kXf = 47;
constexpr std::uint16_t kBeginCellXfs = 617;
constexpr std::uint16_t kEndCellXfs = 618;
constexpr std::uint16_t kBeginStyleXfs = 626;
constexpr std::uint16_t kEndStyleXfs = 627;
constexpr std::uint16_t kFirstCell = 1;
constexpr std::uint16_t kLastCell = 11;
constexpr std::uint16_t kCellRString = 62;
constexpr std::uint16_t kXfApplyBits[] = {0x01, 0x02, 0x04, 0x08, 0x10, 0x20};

template <typename Fn>
void ForEachRecord(const std::string& part, Fn&& fn) {
  ByteSpan cursor{reinterpret_cast<const std::uint8_t*>(part.data()), part.size()};
  while (cursor.size > 0U) {
    auto rec = xlsb::read_record(cursor);
    if (!rec) {
      ADD_FAILURE() << "record: " << rec.error().message;
      return;
    }
    fn(rec.value().type, rec.value().payload);
  }
}

std::string XlsbFlags(ByteSpan payload) {
  if (payload.size < 16U) {
    return "short";
  }
  const unsigned word = payload.data[14] | (payload.data[15] << 8);
  const auto set = [word](int i) { return (word & kXfApplyBits[i]) != 0U; };
  return FlagString(set(0), set(1), set(2), set(3), set(4), set(5));
}

BookFlags ReadXlsb(const Bytes& pkg) {
  BookFlags out;
  std::vector<std::string> cell_xfs;
  bool in_cell_xfs = false;
  bool in_style_xfs = false;
  ForEachRecord(PartText(pkg, "xl/styles.bin"), [&](std::uint16_t type, ByteSpan payload) {
    if (type == kBeginCellXfs || type == kEndCellXfs) {
      in_cell_xfs = type == kBeginCellXfs;
    } else if (type == kBeginStyleXfs || type == kEndStyleXfs) {
      in_style_xfs = type == kBeginStyleXfs;
    } else if (type == kXf && in_cell_xfs) {
      cell_xfs.push_back(XlsbFlags(payload));
    } else if (type == kXf && in_style_xfs) {
      out.style_xfs.insert(XlsbFlags(payload));
    }
  });
  bool found = false;
  ForEachRecord(PartText(pkg, "xl/worksheets/sheet1.bin"), [&](std::uint16_t type, ByteSpan payload) {
    const bool is_cell = (type >= kFirstCell && type <= kLastCell) || type == kCellRString;
    if (!found && is_cell && payload.size >= 8U && xlsb::load_u32(payload.data) == 0U) {
      const std::size_t ixfe = xlsb::load_u32(payload.data + 4) & 0xFFFFFFU;
      if (ixfe < cell_xfs.size()) {
        out.a1 = cell_xfs[ixfe];
      }
      found = true;
    }
  });
  return out;
}

class CellXfsApplyFlagsGolden : public ::testing::TestWithParam<std::string> {
 protected:
  std::string Msg() const { return "case " + GetParam() + " (left: Excel, right: Formulon)"; }
};

TEST_P(CellXfsApplyFlagsGolden, XlsxMatchesExcel) {
  const auto wb = LoadAuthored(GetParam());
  ASSERT_TRUE(wb);
  auto saved = write_ooxml(*wb);
  ASSERT_TRUE(saved) << saved.error().message;
  const BookFlags excel = ReadXlsx(test::read_file_bytes(FixturePath(GetParam(), ".out.xlsx")));
  const BookFlags ours = ReadXlsx(saved.value());
  EXPECT_EQ(excel.a1, ours.a1) << Msg();
  EXPECT_EQ(excel.style_xfs, ours.style_xfs) << Msg();
}

TEST_P(CellXfsApplyFlagsGolden, XlsbMatchesExcel) {
  const auto wb = LoadAuthored(GetParam());
  ASSERT_TRUE(wb);
  auto saved = xlsb::write_xlsb(*wb);
  ASSERT_TRUE(saved) << saved.error().message;
  const BookFlags excel = ReadXlsb(test::read_file_bytes(FixturePath(GetParam(), ".out.xlsb")));
  const BookFlags ours = ReadXlsb(saved.value());
  EXPECT_EQ(excel.a1, ours.a1) << Msg();
  EXPECT_EQ(excel.style_xfs, ours.style_xfs) << Msg();
}

INSTANTIATE_TEST_SUITE_P(Cases, CellXfsApplyFlagsGolden, ::testing::ValuesIn(CaseIds()), CaseName);

}  // namespace
}  // namespace formulon::io
