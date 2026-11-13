//
// The dxf table a save writes, against Excel 16.115 (Mac) re-saving the same
// files. Each fixture gives every dxf user its own solid fill (one factor
// varied per file: active sheet, tab order, CF priority, block split, table
// style element order, a dxf shared across users or equal in content, kinds on
// one sheet, two formats in one pivot, a table colorFilter); the
// expected list is the fill of each `<dxfs>` record Excel wrote, in order.
// Excel writes the same order to xlsx and xlsb.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <map>
#include <regex>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/styles_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "styles.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io {
namespace {

struct Golden {
  const char* file;
  std::vector<std::string> fills;
  /// Records compared in order; the rest are compared as a set (Excel orders
  /// two pivots on one sheet differently from one save to the next).
  std::size_t ordered;
  /// Users the xlsb writer has no record for (custom table style elements,
  /// pivot formats, the sheet colorFilter).
  std::size_t xlsb_unwritten = 6;
};

const std::vector<Golden>& Goldens() {
  static const std::vector<Golden> goldens = {
      {"c0_all",
       {"94D4D3", "967221", "23E1A7", "2E9548", "51CCC8", "6FC658", "DD9A42", "8BC9E0", "56C95B",
        "A2AA3C", "A68D3F", "378278", "4F687F", "5979BC", "D2395E", "476735", "9A3C69", "A94BD0",
        "7B5893", "1E6E22", "88CF51", "7E937B", "2B4291", "4C28CB", "C05F68", "9AB665", "7BCB9B"},
       27},
      {"f10_active_Sheet3",
       {"94D4D3", "967221", "23E1A7", "2E9548", "51CCC8", "6FC658", "DD9A42", "8BC9E0", "56C95B",
        "A2AA3C", "A94BD0", "7B5893", "1E6E22", "88CF51", "7E937B", "2B4291", "A68D3F", "378278",
        "4F687F", "5979BC", "D2395E", "9A3C69", "476735", "4C28CB", "C05F68", "9AB665", "7BCB9B"},
       27},
      {"f2_tabswap34_activeS2",
       {"94D4D3", "967221", "51CCC8", "6FC658", "23E1A7", "2E9548", "DD9A42", "8BC9E0", "56C95B",
        "A2AA3C", "A94BD0", "7B5893", "1E6E22", "88CF51", "7E937B", "2B4291", "476735", "A68D3F",
        "378278", "4F687F", "5979BC", "D2395E", "9A3C69", "4C28CB", "C05F68", "9AB665", "7BCB9B"},
       27},
      {"f3_reverse_activeS2",
       {"56C95B", "A2AA3C", "DD9A42", "8BC9E0", "51CCC8", "6FC658", "23E1A7", "2E9548", "94D4D3",
        "967221", "476735", "A68D3F", "378278", "4F687F", "5979BC", "D2395E", "A94BD0", "7B5893",
        "1E6E22", "88CF51", "7E937B", "2B4291", "9A3C69", "4C28CB", "C05F68", "9AB665", "7BCB9B"},
       27},
      {"f6_priority_reversed",
       {"967221", "94D4D3", "2E9548", "23E1A7", "6FC658", "51CCC8", "8BC9E0", "DD9A42", "A2AA3C",
        "56C95B", "A68D3F", "378278", "4F687F", "5979BC", "D2395E", "476735", "9A3C69", "A94BD0",
        "7B5893", "1E6E22", "88CF51", "7E937B", "2B4291", "4C28CB", "C05F68", "9AB665", "7BCB9B"},
       27},
      {"f7_two_blocks",
       {"94D4D3", "967221", "23E1A7", "2E9548", "51CCC8", "6FC658", "DD9A42", "8BC9E0", "56C95B",
        "A2AA3C", "A68D3F", "378278", "4F687F", "5979BC", "D2395E", "476735", "9A3C69", "A94BD0",
        "7B5893", "1E6E22", "88CF51", "7E937B", "2B4291", "4C28CB", "C05F68", "9AB665", "7BCB9B"},
       27},
      {"f8_elements_header_first",
       {"94D4D3", "967221", "23E1A7", "2E9548", "51CCC8", "6FC658", "DD9A42", "8BC9E0", "56C95B",
        "A2AA3C", "A68D3F", "378278", "4F687F", "5979BC", "D2395E", "476735", "9A3C69", "A94BD0",
        "7B5893", "1E6E22", "88CF51", "7E937B", "2B4291", "C05F68", "4C28CB", "7BCB9B", "9AB665"},
       27},
      {"f9_cf_table",
       {"94D4D3", "9AB665", "88CF51", "2E9548", "51CCC8", "4F687F", "DD9A42", "8BC9E0", "56C95B",
        "C05F68", "A68D3F", "A2AA3C", "6FC658", "4F687F", "A94BD0", "4C28CB", "5979BC", "23E1A7",
        "7B5893", "476735", "9A3C69", "967221", "7BCB9B", "378278", "7E937B", "1E6E22", "D2395E"},
       27},
      {"f9_style_pivot",
       {"94D4D3", "9AB665", "88CF51", "2E9548", "51CCC8", "4F687F", "DD9A42", "8BC9E0", "56C95B",
        "C05F68", "7BCB9B", "A68D3F", "A2AA3C", "6FC658", "A94BD0", "4C28CB", "1E6E22", "D2395E",
        "23E1A7", "476735", "7B5893", "9A3C69", "967221", "5979BC", "378278", "7E937B", "1E6E22"},
       27},
      {"f9_cf_cf_same_sheet",
       {"9AB665", "88CF51", "2E9548", "51CCC8", "4F687F", "4F687F", "DD9A42", "8BC9E0", "56C95B",
        "C05F68", "A68D3F", "A2AA3C", "6FC658", "A94BD0", "94D4D3", "4C28CB", "5979BC", "23E1A7",
        "7B5893", "476735", "9A3C69", "967221", "7BCB9B", "378278", "7E937B", "1E6E22", "D2395E"},
       27},
      {"r5_cf_same_content",
       {"94D4D3", "967221", "23E1A7", "2E9548", "6FC658", "6FC658", "DD9A42", "8BC9E0", "56C95B",
        "A2AA3C", "A68D3F", "378278", "4F687F", "5979BC", "D2395E", "476735", "9A3C69", "A94BD0",
        "7B5893", "1E6E22", "88CF51", "7E937B", "2B4291", "4C28CB", "C05F68", "9AB665", "7BCB9B"},
       27},
      {"r2_pivot_two_formats",
       {"23E1A7", "2B4291", "7B5893", "56C95B", "A68D3F", "5979BC", "88CF51", "4C28CB", "7BCB9B", "4F687F",
        "476735", "95D555", "6FC658", "9A3C69", "51CCC8", "7E937B", "C05F68", "94D4D3", "DD9A42", "D2395E",
        "1E6E22", "8BC9E0", "9AB665", "C4D431", "378278", "A2AA3C", "967221", "2E9548"},
       28,
       7},
      {"r2_pivot_two_formats_rev",
       {"23E1A7", "2B4291", "7B5893", "56C95B", "A68D3F", "5979BC", "88CF51", "4C28CB", "7BCB9B", "4F687F",
        "476735", "95D555", "6FC658", "9A3C69", "51CCC8", "7E937B", "94D4D3", "C05F68", "DD9A42", "D2395E",
        "1E6E22", "8BC9E0", "9AB665", "C4D431", "378278", "A2AA3C", "967221", "2E9548"},
       28,
       7},
      {"r4_table_colorfilter",
       {"23E1A7", "2B4291", "7B5893", "56C95B", "A68D3F", "5979BC", "88CF51", "4C28CB", "7BCB9B", "4F687F",
        "476735", "95D555", "6FC658", "9A3C69", "51CCC8", "7E937B", "94D4D3", "DD9A42", "D2395E", "1E6E22",
        "C05F68", "8BC9E0", "9AB665", "C4D431", "378278", "A2AA3C", "967221", "2E9548"},
       28},
      {"f9_cf_cf_two_sheets",
       {"9AB665", "C05F68", "88CF51", "2E9548", "51CCC8", "4F687F", "DD9A42", "8BC9E0", "56C95B",
        "C05F68", "A68D3F", "A2AA3C", "6FC658", "A94BD0", "94D4D3", "4C28CB", "5979BC", "23E1A7",
        "7B5893", "476735", "9A3C69", "967221", "7BCB9B", "378278", "7E937B", "1E6E22", "D2395E"},
       27},
      // Authored blocks K2, H9, I1 save as H9, I1, K2; the fills follow the saved order.
      {"cf_unsorted_blocks", {"22CC22", "2222CC", "CC2222"}, 3, 0},
      {"f1_one_sheet_cf", {"4C28CB", "A68D3F", "7B5893", "378278", "2E9548", "8BC9E0", "51CCC8", "7BCB9B"}, 8},
      {"f5_two_tables_two_pivots", {"51CCC8", "7BCB9B", "8BC9E0", "7B5893", "378278", "2E9548", "A68D3F", "4C28CB"}, 6},
  };
  return goldens;
}

/// The solid fill colour (RRGGBB) of each dxf, in table order.
std::vector<std::string> Fills(const Workbook& wb) {
  static const std::regex kFg("<fgColor rgb=\"[0-9A-Fa-f]{2}([0-9A-Fa-f]{6})\"");
  std::vector<std::string> out;
  for (const DifferentialFormat& dxf : wb.styles().dxfs) {
    const std::string body = dxf_body_fragment(dxf);
    std::smatch m;
    out.push_back(std::regex_search(body, m, kFg) ? m[1].str() : std::string("?"));
  }
  return out;
}

Workbook LoadFixture(const char* name) {
  const std::vector<std::uint8_t> bytes =
      test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/dxf_order/" + name + ".xlsx");
  auto read = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read)) << name;
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

void ExpectGolden(const Golden& golden, const std::vector<std::string>& saved, const char* format) {
  ASSERT_EQ(saved.size(), golden.fills.size()) << golden.file << " " << format;
  for (std::size_t i = 0; i < golden.ordered; ++i) {
    EXPECT_EQ(saved[i], golden.fills[i]) << golden.file << " " << format << " record " << i;
  }
  std::vector<std::string> rest(saved.begin() + static_cast<std::ptrdiff_t>(golden.ordered), saved.end());
  std::vector<std::string> expected(golden.fills.begin() + static_cast<std::ptrdiff_t>(golden.ordered),
                                    golden.fills.end());
  std::sort(rest.begin(), rest.end());
  std::sort(expected.begin(), expected.end());
  EXPECT_EQ(rest, expected) << golden.file << " " << format;
}

TEST(DxfOrderGolden, XlsxSaveWritesExcelsOrder) {
  for (const Golden& golden : Goldens()) {
    const Workbook wb = LoadFixture(golden.file);
    auto bytes = write_ooxml(wb);
    ASSERT_TRUE(static_cast<bool>(bytes)) << golden.file;
    auto reread = read_ooxml(ByteSpan{bytes.value().data(), bytes.value().size()});
    ASSERT_TRUE(static_cast<bool>(reread)) << golden.file;
    ExpectGolden(golden, Fills(reread.value().workbook), "xlsx");
  }
}

/// `golden` restricted to the records `saved` holds, keeping Excel's order.
/// The xlsb writer does not write custom table styles, pivot tables or a
/// sheet autoFilter from the model, so those users have no record there.
Golden WrittenPart(const Golden& golden, const std::vector<std::string>& saved) {
  std::map<std::string, std::size_t> left;
  for (const std::string& fill : saved) {
    ++left[fill];
  }
  Golden out{golden.file, {}, 0, golden.xlsb_unwritten};
  for (std::size_t i = 0; i < golden.fills.size(); ++i) {
    if (left[golden.fills[i]] > 0U) {
      --left[golden.fills[i]];
      out.fills.push_back(golden.fills[i]);
      out.ordered += i < golden.ordered ? 1U : 0U;
    }
  }
  return out;
}

TEST(DxfOrderGolden, XlsbSaveWritesExcelsOrder) {
  for (const Golden& golden : Goldens()) {
    const Workbook wb = LoadFixture(golden.file);
    auto bytes = xlsb::write_xlsb(wb);
    ASSERT_TRUE(static_cast<bool>(bytes)) << golden.file;
    auto reread = xlsb::read_xlsb(ByteSpan{bytes.value().data(), bytes.value().size()});
    ASSERT_TRUE(static_cast<bool>(reread)) << golden.file;
    const std::vector<std::string> saved = Fills(reread.value().workbook);
    const Golden written = WrittenPart(golden, saved);
    // Every CF and table user is written in xlsb.
    EXPECT_GE(written.fills.size(), golden.fills.size() - golden.xlsb_unwritten) << golden.file;
    ExpectGolden(written, saved, "xlsb");
  }
}

TEST(DxfOrderGolden, SaveDoesNotTouchTheLiveTable) {
  const Workbook wb = LoadFixture("f9_cf_table");
  const std::vector<std::string> before = Fills(wb);
  ASSERT_TRUE(static_cast<bool>(write_ooxml(wb)));
  ASSERT_TRUE(static_cast<bool>(xlsb::write_xlsb(wb)));
  EXPECT_EQ(Fills(wb), before);
}

}  // namespace
}  // namespace formulon::io
