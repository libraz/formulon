#include "chart/chart_source.h"

#include <algorithm>
#include <string>

#include "defined_name.h"
#include "eval/builtin_names.h"
#include "eval/defined_name_resolve.h"
#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "eval/range_resolvers.h"
#include "parser/ast.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "utils/arena.h"
#include "utils/number_text.h"
#include "utils/resource_budget.h"
#include "utils/strings.h"
#include "workbook.h"

namespace formulon {
namespace {

Error invalid(std::string_view formula, std::string_view why) {
  return make_error(FormulonErrorCode::kIoChartInvalidReference, "chart: invalid reference",
                    std::string(formula) + ": " + std::string(why));
}

struct Resolver {
  Resolver(const Workbook& w, std::size_t host, const eval::FunctionRegistry& reg, std::string_view f)
      : wb(w), host_sheet(host), registry(reg), formula(f) {}

  const Workbook& wb;
  std::size_t host_sheet;
  const eval::FunctionRegistry& registry;
  std::string_view formula;
  Arena arena{/*initial_chunk_bytes=*/4096, kMaxEvalArenaBytes};
  std::vector<SheetRect> areas;

  /// Maps a resolver sheet qualifier (empty = bound sheet) to a sheet index.
  Expected<std::size_t, Error> sheet_of(std::string_view qualifier) const {
    if (qualifier.empty()) {
      if (host_sheet == kChartNoHostSheet) {
        return invalid(formula, "unqualified reference on a chartsheet");
      }
      return host_sheet;
    }
    const std::size_t index = wb.sheet_index_by_name(qualifier);
    if (index >= wb.sheet_count()) {
      return invalid(formula, "unknown sheet");
    }
    return index;
  }

  Expected<void, Error> push(std::string_view qualifier, std::uint32_t top, std::uint32_t left, std::uint32_t bottom,
                             std::uint32_t right) {
    auto sheet = sheet_of(qualifier);
    if (!sheet) {
      return sheet.error();
    }
    areas.push_back(SheetRect{sheet.value(), top, left, bottom, right});
    return Expected<void, Error>::Ok();
  }

  /// Bare `Name`: an active name scope takes precedence; a root chartsheet
  /// without one sees workbook-scoped names only.
  const DefinedName* find_bare_name(std::string_view name, const eval::EvalContext& ctx) const {
    if (host_sheet != kChartNoHostSheet || ctx.name_scope_sheet() >= 0) {
      return eval::find_defined_name(ctx, name);
    }
    for (const DefinedName& def : wb.defined_names()) {
      if (def.local_sheet_id < 0 && strings::case_insensitive_eq(def.name, name)) {
        return &def;
      }
    }
    return nullptr;
  }

  Expected<void, Error> resolve_name(const parser::AstNode& node, const eval::EvalContext& ctx) {
    const DefinedName* def = nullptr;
    if (parser::is_self_book_name_ref(node)) {
      def = eval::find_self_book_defined_name(wb, node.as_external_ref_name());
    } else if (!node.as_name_sheet().empty()) {
      def = eval::find_sheet_defined_name(wb, node.as_name_sheet(), node.as_name());
    } else {
      def = find_bare_name(node.as_name(), ctx);
    }
    if (def == nullptr) {
      return invalid(formula, "undefined name");
    }
    eval::DefinedNameFrame frame;
    eval::EvalContext body_ctx = ctx;
    ErrorCode err = ErrorCode::Ref;
    const parser::AstNode* body = eval::prepare_defined_name_body(def, arena, ctx, &frame, &body_ctx, &err);
    if (body == nullptr) {
      return invalid(formula, "unusable name body");
    }
    return resolve(*body, body_ctx);
  }

  Expected<void, Error> resolve_spill(const parser::AstNode& node, const eval::EvalContext& ctx) {
    std::string_view qualifier;
    std::uint32_t row = 0;
    std::uint32_t col = 0;
    if (const parser::AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
      std::uint32_t bottom = 0;
      std::uint32_t right = 0;
      ErrorCode err = ErrorCode::Ref;
      if (!eval::resolve_reference_rect(*anchor, arena, registry, ctx, &qualifier, &row, &col, &bottom, &right, &err)) {
        return invalid(formula, "spill anchor is not a reference");
      }
    } else {
      const parser::Reference& ref = node.as_spill_ref();
      qualifier = ref.sheet;
      row = ref.row;
      col = ref.col;
    }
    auto sheet = sheet_of(qualifier);
    if (!sheet) {
      return sheet.error();
    }
    const SpillRegion* region = wb.sheet(sheet.value()).spill_region_at_anchor(row, col);
    if (region == nullptr || region->rows == 0 || region->cols == 0) {
      return invalid(formula, "no spill at anchor");
    }
    areas.push_back(SheetRect{sheet.value(), row, col, row + region->rows - 1, col + region->cols - 1});
    return Expected<void, Error>::Ok();
  }

  Expected<void, Error> resolve(const parser::AstNode& node, const eval::EvalContext& ctx) {
    switch (node.kind()) {
      case parser::NodeKind::UnionOp:
        for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
          auto child = resolve(node.as_union_child(i), ctx);
          if (!child) {
            return child;
          }
        }
        return Expected<void, Error>::Ok();
      case parser::NodeKind::NameRef:
        return resolve_name(node, ctx);
      case parser::NodeKind::ExternalRef:
        if (parser::is_self_book_name_ref(node)) {
          return resolve_name(node, ctx);
        }
        return invalid(formula, "external reference");
      case parser::NodeKind::SpillRef:
        return resolve_spill(node, ctx);
      case parser::NodeKind::Ref:
      case parser::NodeKind::RangeOp:
      case parser::NodeKind::IntersectOp:
      case parser::NodeKind::Call:
        break;
      default:
        return invalid(formula, "not a reference");
    }
    std::string_view qualifier;
    std::uint32_t top = 0;
    std::uint32_t left = 0;
    std::uint32_t bottom = 0;
    std::uint32_t right = 0;
    ErrorCode err = ErrorCode::Ref;
    if (!eval::resolve_reference_rect(node, arena, registry, ctx, &qualifier, &top, &left, &bottom, &right, &err)) {
      return invalid(formula, "not a reference");
    }
    return push(qualifier, top, left, bottom, right);
  }
};

}  // namespace

Expected<ResolvedSource, Error> resolve_chart_source(const Workbook& wb, std::size_t host_sheet,
                                                     std::string_view formula, const eval::FunctionRegistry& registry) {
  if (!formula.empty() && formula.front() == '=') {
    formula.remove_prefix(1);
  }
  if (wb.sheet_count() == 0 || (host_sheet != kChartNoHostSheet && host_sheet >= wb.sheet_count())) {
    return invalid(formula, "no such host sheet");
  }
  eval::ScopedEvalProfile profile_scope(wb.excel_profile());
  Resolver resolver(wb, host_sheet, registry, formula);
  const parser::AstNode* root = eval::parse_formula_entry(formula, resolver.arena);
  if (root == nullptr) {
    return invalid(formula, "unparseable");
  }
  const Sheet& bound = wb.sheet(host_sheet == kChartNoHostSheet ? 0U : host_sheet);
  const eval::EvalContext ctx = eval::EvalContext::workbook_only(wb, bound);
  auto resolved = resolver.resolve(*root, ctx);
  if (!resolved) {
    return resolved.error();
  }
  ResolvedSource out;
  out.areas = std::move(resolver.areas);
  return out;
}

SheetRect trim_to_populated(const Workbook& wb, SheetRect r) {
  const bool whole_col = r.top == 0 && r.bottom == Sheet::kMaxRows - 1;
  const bool whole_row = r.left == 0 && r.right == Sheet::kMaxCols - 1;
  if (!whole_col && !whole_row) {
    return r;
  }
  const auto extent = wb.sheet(r.sheet).populated_extent(0, 0, Sheet::kMaxRows - 1, Sheet::kMaxCols - 1);
  if (!extent.has_value()) {
    r.top = 1;
    r.bottom = 0;
    return r;
  }
  if (whole_col) {
    r.bottom = std::min(r.bottom, extent->last_row);
  }
  if (whole_row) {
    r.right = std::min(r.right, extent->last_col);
  }
  return r;
}

bool row_hidden(const Sheet& sheet, std::uint32_t row) {
  if (const RowLayout* entry = find_row_override(sheet.layout(), row); entry != nullptr) {
    return entry->hidden;
  }
  return sheet.format_defaults().zero_height;
}

bool col_hidden(const Sheet& sheet, std::uint32_t col) {
  for (const ColumnLayout& span : sheet.layout().columns) {
    if (col >= span.first && col <= span.last) {
      return span.hidden;
    }
  }
  return false;
}

std::string count_text(std::uint64_t n) {
  char buf[24];
  const int len = format_unsigned(buf, sizeof(buf), n);
  return std::string(buf, len > 0 ? static_cast<std::size_t>(len) : 0U);
}

}  // namespace formulon
