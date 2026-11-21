// @size-budget: 56 KB

#include "io/xlsb/feature_formula.h"

#include <algorithm>
#include <string>
#include <utility>

#include "io/cf_reader.h"
#include "io/xlsb/ptg_targets.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/tail_refs.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"

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
    const MergeRange r = read_rfx(cursor).value();
    if (!in_grid(r)) {
      return false;
    }
    out.push_back(r);
  }
  return true;
}

void emit_sqref(std::vector<std::uint8_t>& dst, const std::vector<MergeRange>& ranges) {
  emit_u32(dst, static_cast<std::uint32_t>(ranges.size()));
  for (const MergeRange& r : ranges) {
    emit_rfx(dst, r);
  }
}

FeatureFormulaLoadView::FeatureFormulaLoadView(const FeatureFormulaWriteContext& ctx,
                                               const std::vector<XlsbExternSheetEntry>& extern_sheets,
                                               std::string_view own_sheet)
    : context_{ctx.sheet_names, names_, sheet_ranges_, no_books_,
               static_cast<std::size_t>(std::find(ctx.sheet_names.begin(), ctx.sheet_names.end(), own_sheet) -
                                        ctx.sheet_names.begin())} {
  for (const XlsbExternSheetEntry& entry : extern_sheets) {
    const auto itab = [&ctx](const std::string& name) {
      const auto at = std::find(ctx.sheet_names.begin(), ctx.sheet_names.end(), name);
      return at == ctx.sheet_names.end() ? -1 : static_cast<std::int32_t>(at - ctx.sheet_names.begin());
    };
    XlsbSheetRange range;
    if (entry.deleted_sheet) {
      range.itab_first = range.itab_last = kXtiDeletedSheet;
    } else if (entry.external_book != 0U || entry.unresolved) {
      range.external_book = entry.external_book != 0U ? entry.external_book : 1U;
    } else if (entry.first.empty() && entry.last.empty()) {
      range.itab_first = range.itab_last = -2;
    } else {
      range.itab_first = itab(entry.first);
      range.itab_last = itab(entry.last);
    }
    sheet_ranges_.push_back(range);
  }
  for (const auto& [key, ilbl] : ctx.name_table) {
    if (ilbl != 0U) {
      names_.resize(std::max<std::size_t>(names_.size(), ilbl));
      names_[ilbl - 1U].name = key;
    }
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
  ASSIGN_OR_RETURN(const std::uint32_t cce, read_u32(cursor));
  if (cce > cursor.size) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "feature formula rgce truncated",
                      "context=xlsb_feature_formula");
  }
  std::vector<std::uint8_t> rgce(cursor.data, cursor.data + cce);
  cursor.data += cce;
  cursor.size -= cce;
  ASSIGN_OR_RETURN(const std::uint32_t cb, read_u32(cursor));
  if (cb > cursor.size) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "feature formula rgcb truncated",
                      "context=xlsb_feature_formula");
  }
  const ByteSpan rgcb{cursor.data, cb};
  cursor.data += cb;
  cursor.size -= cb;
  if (rgce.empty()) {
    return std::string();
  }
  Arena arena(kDefaultArenaChunkBytes, kMaxLoadArenaBytes);
  ASSIGN_OR_RETURN(
      const auto ast,
      decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, rgcb, arena, ctx.sheet_names, ctx.name_table, ctx.sheet_ranges,
                  ctx.external_books, static_cast<std::int32_t>(ctx.sheet_index), PtgBaseCell{base_row, base_col}));
  return canonical_feature_formula(parser::format_formula(*ast));
}

Expected<EncodedFeatureFormula, Error> encode_feature_formula(std::string_view text, std::uint32_t base_row,
                                                              std::uint32_t base_col, PtgRootClass root_class,
                                                              const FeatureFormulaWriteContext& ctx,
                                                              PtgEvaluation evaluation) {
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
  ASSIGN_OR_RETURN(const auto encoded, encode_ptgs(*root, ctx.sheet_names, ctx.sheet_ranges, ctx.name_table, root_class,
                                                   PtgBaseCell{base_row, base_col}, evaluation));
  const std::vector<std::uint8_t>& rgce = encoded.rgce;
  out.cb_fmla =
      legacy_cf_formula_size_field(rgce.data(), rgce.size()).value_or(static_cast<std::uint32_t>(rgce.size()));
  const std::vector<std::uint8_t>& rgcb = encoded.rgcb;
  emit_u32(out.bytes, static_cast<std::uint32_t>(rgce.size()));
  out.bytes.insert(out.bytes.end(), rgce.begin(), rgce.end());
  emit_u32(out.bytes, static_cast<std::uint32_t>(rgcb.size()));
  out.bytes.insert(out.bytes.end(), rgcb.begin(), rgcb.end());
  return out;
}

Expected<EncodedFeatureFormula, Error> encode_cf_formula(std::string_view text, const MergeRange& base,
                                                         const FeatureFormulaWriteContext& ctx) {
  return encode_feature_formula(text, base.first_row, base.first_col, PtgRootClass::kValue, ctx,
                                PtgEvaluation::kConditionalFormat);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
