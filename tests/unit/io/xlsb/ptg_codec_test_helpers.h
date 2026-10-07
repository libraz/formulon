#pragma once

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/ptg_writer.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "utils/arena.h"

namespace formulon::io::xlsb::ptg_codec_test_support {

// Parses `formula` (without leading `=`), encodes to Ptg, decodes back,
// and returns the re-formatted formula text. `sheet_names` resolves a
// qualified reference's sheet to its 0-based index on both sides; the
// `SheetRangeTable` / `XlsbSheetRange` list that actually carries the
// `ixti` numbering is built here (via `collect_ptg_sheet_ranges` on the
// encode side, mirrored 1:1 into `XlsbSheetRange`s for decode) so a
// single-sheet qualified reference and a genuine 3-D range share one
// `ixti` space exactly as the production writer does.
inline std::string RoundTrip(std::string_view formula, const std::vector<std::string>& sheet_names = {}) {
  Arena enc_arena;
  parser::Parser p(formula, enc_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty()) << "parse errors for: " << formula;

  SheetRangeTable sheet_ranges;
  std::unordered_set<std::uint64_t> seen;
  collect_ptg_sheet_ranges(*root, sheet_names, sheet_ranges, seen);

  auto encoded = encode_ptgs(*root, sheet_names, sheet_ranges, {}, PtgRootClass::kValue);
  EXPECT_TRUE(static_cast<bool>(encoded))
      << "encode failed for: " << formula << " | " << (encoded ? "" : encoded.error().message);
  if (!encoded) {
    return "<encode-failed>";
  }

  std::vector<XlsbSheetRange> decode_ranges;
  decode_ranges.reserve(sheet_ranges.xti.size());
  for (const XtiEntry& entry : sheet_ranges.xti) {
    decode_ranges.push_back(XlsbSheetRange{entry.first, entry.last});
  }

  Arena dec_arena;
  ByteSpan rgce{encoded.value().rgce.data(), encoded.value().rgce.size()};
  ByteSpan rgcb{encoded.value().rgcb.data(), encoded.value().rgcb.size()};
  auto decoded = decode_ptgs(rgce, rgcb, dec_arena, sheet_names, {}, decode_ranges, {});
  EXPECT_TRUE(static_cast<bool>(decoded))
      << "decode failed for: " << formula << " | " << (decoded ? "" : decoded.error().message);
  if (!decoded) {
    return "<decode-failed>";
  }
  return parser::format_formula(*decoded.value());
}
inline EncodedFormula EncodeOnSheet1(std::string_view formula, PtgRootClass root_class) {
  const std::vector<std::string> sheets = {"Sheet1"};
  Arena arena;
  parser::Parser p(formula, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << formula;
  EXPECT_TRUE(p.errors().empty()) << formula;
  if (root == nullptr) {
    return {};
  }
  SheetRangeTable ranges;
  std::unordered_set<std::uint64_t> seen;
  collect_ptg_sheet_ranges(*root, sheets, ranges, seen);
  // A cell formula is encoded as typed into Excel 365: with the dynamic-array
  // mark exactly where Excel gives it one.
  const PtgEvaluation evaluation =
      root_class == PtgRootClass::kReference || formula_is_dynamic_array(*root, NameShapes())
          ? PtgEvaluation::kDynamicArray
          : PtgEvaluation::kLegacy;
  auto encoded = encode_ptgs(*root, sheets, ranges, {}, root_class, std::nullopt, evaluation);
  EXPECT_TRUE(static_cast<bool>(encoded)) << formula << " | " << (encoded ? "" : encoded.error().message);
  return encoded ? encoded.value() : EncodedFormula{};
}

// Bytes as Excel 365 saved each formula. A
// union or intersection of plain references in a cell sits behind the memory
// token carrying Excel's precomputed result; `PtgMemArea`'s leading four

}  // namespace formulon::io::xlsb::ptg_codec_test_support
