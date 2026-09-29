#include "io/xlsb/feature_formula.h"

#include <string>
#include <utility>

#include "io/cf_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/resource_budget.h"

namespace formulon {
namespace io {
namespace xlsb {

bool read_sqref(ByteSpan& cursor, std::vector<MergeRange>& out) {
  auto count = read_u32(cursor);
  if (!count || count.value() == 0U || count.value() > cursor.size / 16U) {
    return false;
  }
  out.clear();
  for (std::uint32_t i = 0; i < count.value(); ++i) {
    MergeRange r;
    r.first_row = read_u32(cursor).value();
    r.last_row = read_u32(cursor).value();
    r.first_col = read_u32(cursor).value();
    r.last_col = read_u32(cursor).value();
    if (r.last_row >= Sheet::kMaxRows || r.last_col >= Sheet::kMaxCols || r.first_row > r.last_row ||
        r.first_col > r.last_col) {
      return false;
    }
    out.push_back(r);
  }
  return true;
}

void emit_sqref(std::vector<std::uint8_t>& dst, const std::vector<MergeRange>& ranges) {
  emit_u32(dst, static_cast<std::uint32_t>(ranges.size()));
  for (const MergeRange& r : ranges) {
    emit_u32(dst, r.first_row);
    emit_u32(dst, r.last_row);
    emit_u32(dst, r.first_col);
    emit_u32(dst, r.last_col);
  }
}

MergeRange sqref_base(const std::vector<MergeRange>& ranges) {
  MergeRange base{Sheet::kMaxRows, Sheet::kMaxCols, 0, 0};
  for (const MergeRange& r : ranges) {
    base.first_row = r.first_row < base.first_row ? r.first_row : base.first_row;
    base.first_col = r.first_col < base.first_col ? r.first_col : base.first_col;
  }
  return base;
}

Expected<std::string, Error> read_feature_formula(ByteSpan& cursor, std::uint32_t base_row, std::uint32_t base_col,
                                                  const FeatureFormulaReadContext& ctx) {
  auto cce = read_u32(cursor);
  if (!cce) {
    return cce.error();
  }
  if (cce.value() > cursor.size) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "feature formula rgce truncated",
                      "context=xlsb_feature_formula");
  }
  std::vector<std::uint8_t> rgce(cursor.data, cursor.data + cce.value());
  cursor.data += cce.value();
  cursor.size -= cce.value();
  auto cb = read_u32(cursor);
  if (!cb) {
    return cb.error();
  }
  if (cb.value() > cursor.size) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "feature formula rgcb truncated",
                      "context=xlsb_feature_formula");
  }
  const ByteSpan rgcb{cursor.data, cb.value()};
  cursor.data += cb.value();
  cursor.size -= cb.value();
  if (rgce.empty()) {
    return std::string();
  }
  Arena arena(/*initial_chunk_bytes=*/4096, kMaxLoadArenaBytes);
  auto ast =
      decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, rgcb, arena, ctx.sheet_names, ctx.name_table, ctx.sheet_ranges,
                  ctx.external_books, static_cast<std::int32_t>(ctx.sheet_index), PtgBaseCell{base_row, base_col});
  if (!ast) {
    return ast.error();
  }
  return canonical_feature_formula(parser::format_formula(*ast.value()));
}

Expected<EncodedFeatureFormula, Error> encode_feature_formula(std::string_view text, std::uint32_t base_row,
                                                              std::uint32_t base_col, PtgRootClass root_class,
                                                              const FeatureFormulaWriteContext& ctx) {
  EncodedFeatureFormula out;
  if (text.empty()) {
    emit_u32(out.bytes, 0U);
    emit_u32(out.bytes, 0U);
    return out;
  }
  Arena arena;
  parser::Parser parser(text, arena);
  parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "feature formula failed to parse",
                      "context=xlsb_feature_formula formula=" + std::string(text));
  }
  auto encoded = encode_ptgs(*root, ctx.sheet_names, ctx.sheet_ranges, ctx.name_table, root_class,
                             PtgBaseCell{base_row, base_col});
  if (!encoded) {
    return encoded.error();
  }
  const std::vector<std::uint8_t>& rgce = encoded.value().rgce;
  out.cb_fmla = static_cast<std::uint32_t>(rgce.size());
  const std::vector<std::uint8_t>& rgcb = encoded.value().rgcb;
  emit_u32(out.bytes, static_cast<std::uint32_t>(rgce.size()));
  out.bytes.insert(out.bytes.end(), rgce.begin(), rgce.end());
  emit_u32(out.bytes, static_cast<std::uint32_t>(rgcb.size()));
  out.bytes.insert(out.bytes.end(), rgcb.begin(), rgcb.end());
  return out;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
