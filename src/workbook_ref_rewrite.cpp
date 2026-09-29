#include "workbook_ref_rewrite.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "defined_name.h"
#include "eval/dep_graph.h"
#include "eval/recalc_engine.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/parser.h"
#include "pivot/pivot_cache.h"
#include "sheet.h"
#include "sheet_name.h"
#include "table.h"
#include "utils/arena.h"

namespace formulon {
namespace {

// Rewrites a single formula text through `transform`. Returns the
// rewritten body alongside a flag indicating whether any reference was
// actually changed; an unchanged body lets the caller skip the
// dep-graph re-register and avoid touching the cell.
struct FormulaRewriteResult {
  std::string text;
  bool changed = false;
  bool error = false;  // Parse failure or arena exhaustion.
};

FormulaRewriteResult rewrite_formula(std::string_view formula, const parser::RefTransform& transform) {
  FormulaRewriteResult out;
  out.text.assign(formula);
  std::string_view body = formula;
  bool had_equals = false;
  if (!body.empty() && body.front() == '=') {
    body = body.substr(1);
    had_equals = true;
  }
  if (body.empty()) {
    return out;
  }
  Arena arena;
  parser::Parser parser(body, arena);
  parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    out.error = true;
    return out;
  }
  const parser::AstNode* shifted = parser::shift_refs(*root, arena, transform);
  if (shifted == nullptr) {
    out.error = true;
    return out;
  }
  if (shifted == root) {
    return out;  // Identity walk; no change required.
  }
  std::string rewritten;
  if (had_equals) {
    rewritten.push_back('=');
  }
  rewritten.append(parser::format_formula(*shifted));
  out.text = std::move(rewritten);
  out.changed = true;
  return out;
}

}  // namespace

void rewrite_sheet_metadata_formulas(std::vector<Sheet>& sheets,
                                     const std::vector<const parser::RefTransform*>& per_sheet,
                                     std::vector<TableMetadata>& tables,
                                     std::vector<std::unique_ptr<pivot::PivotCache>>& pivot_caches,
                                     std::string_view direct_sheet_old, std::string_view direct_sheet_new,
                                     std::string_view removed_sheet_name,
                                     std::vector<std::uint32_t>& dropped_cache_ids) {
  dropped_cache_ids.clear();
  if (per_sheet.empty()) {
    // No transform to apply at all. A short `per_sheet` is instead absorbed
    // per sheet below: degrading one sheet to the first slot keeps the rest
    // rewritten, whereas returning here would silently rewrite nothing and
    // leave every holder pointing at pre-edit coordinates.
    return;
  }

  const auto rewrite_field = [](std::string& field, const parser::RefTransform& transform) {
    const FormulaRewriteResult result = rewrite_formula(field, transform);
    if (!result.changed) {
      return false;
    }
    field = result.text;
    return true;
  };

  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    Sheet& sheet = sheets[sheet_idx];
    const parser::RefTransform& transform = *per_sheet[sheet_idx < per_sheet.size() ? sheet_idx : 0U];
    for (cf::ConditionalFormat& conditional_format : sheet.mutable_conditional_formats()) {
      for (cf::CFRule& rule : conditional_format.rules) {
        if (rule.formula1.has_value()) {
          rewrite_field(*rule.formula1, transform);
        }
        if (rule.formula2.has_value()) {
          rewrite_field(*rule.formula2, transform);
        }
        const auto rewrite_cfvo = [&rewrite_field, &transform](cf::CfValueObject& value) {
          if (value.type == cf::CfvoType::Formula) {
            rewrite_field(value.value, transform);
          }
        };
        if (rule.color_scale.has_value()) {
          for (cf::CfValueObject& value : rule.color_scale->thresholds) {
            rewrite_cfvo(value);
          }
        }
        if (rule.icon_set.has_value()) {
          rewrite_cfvo(rule.icon_set->floor);
          for (cf::CfValueObject& value : rule.icon_set->thresholds) {
            rewrite_cfvo(value);
          }
        }
        if (rule.data_bar.has_value()) {
          rewrite_cfvo(rule.data_bar->min);
          rewrite_cfvo(rule.data_bar->max);
        }
      }
    }

    for (Hyperlink& hyperlink : sheet.mutable_hyperlinks()) {
      if (hyperlink.location.empty()) {
        continue;
      }
      const bool has_fragment_prefix = hyperlink.location.front() == '#';
      const std::string_view body =
          has_fragment_prefix ? std::string_view(hyperlink.location).substr(1) : std::string_view(hyperlink.location);
      const FormulaRewriteResult result = rewrite_formula(body, transform);
      if (!result.changed) {
        continue;
      }
      // A fragment marker is a transport prefix, not part of the formula.
      // Avoid turning the removal result `#REF!` into the invalid `##REF!`.
      if (has_fragment_prefix && result.text == "#REF!") {
        hyperlink.location = "#REF!";
      } else {
        std::string rewritten;
        rewritten.reserve(result.text.size() + (has_fragment_prefix ? 1U : 0U));
        if (has_fragment_prefix) {
          rewritten.push_back('#');
        }
        rewritten.append(result.text);
        hyperlink.location = std::move(rewritten);
      }
    }

    for (DataValidation& validation : sheet.mutable_validations()) {
      rewrite_field(validation.formula1, transform);
      rewrite_field(validation.formula2, transform);
    }
  }

  for (TableMetadata& table : tables) {
    const parser::RefTransform& transform = *per_sheet[table.sheet_index < per_sheet.size() ? table.sheet_index : 0U];
    rewrite_field(table.ref, transform);
    for (TableColumn& column : table.columns) {
      rewrite_field(column.calculated_column_formula, transform);
    }
  }

  for (std::unique_ptr<pivot::PivotCache>& cache : pivot_caches) {
    if (cache == nullptr) {
      continue;
    }
    pivot::WorksheetSource& source = cache->mutable_worksheet_source();
    if (!direct_sheet_old.empty() && sheet_names::equal(source.sheet, direct_sheet_old)) {
      source.sheet.assign(direct_sheet_new);
    }
    std::size_t owner_sheet = 0;
    for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
      if (sheet_names::equal(sheets[sheet_idx].name(), source.sheet)) {
        owner_sheet = sheet_idx;
        break;
      }
    }
    const FormulaRewriteResult ref_result = rewrite_formula(source.ref, *per_sheet[owner_sheet]);
    if (ref_result.changed) {
      source.ref = ref_result.text;
    }
    if (!removed_sheet_name.empty() && source.present &&
        (sheet_names::equal(source.sheet, removed_sheet_name) || ref_result.changed)) {
      dropped_cache_ids.push_back(cache->cache_id());
    }
  }
}

void rewrite_workbook_references(std::vector<Sheet>& sheets, std::vector<DefinedName>& defined_names,
                                 std::vector<TableMetadata>& tables,
                                 std::vector<std::unique_ptr<pivot::PivotCache>>& pivot_caches,
                                 const parser::RefTransform& transform,
                                 const eval::RecalcEngine::LockedMutator& mutator, std::string_view direct_sheet_old,
                                 std::string_view direct_sheet_new, std::string_view removed_sheet_name,
                                 std::vector<std::uint32_t>& dropped_cache_ids, bool& defined_names_changed) {
  defined_names_changed = false;

  struct CellUpdate {
    std::size_t sheet_index = 0;
    std::uint32_t row = 0;
    std::uint32_t col = 0;
    std::string formula;
  };
  std::vector<CellUpdate> cell_updates;

  // Collect first so replacing a formula cannot invalidate the row/cell
  // references used by the scan. The cell store's physical coordinates do
  // not move during a sheet rename/removal transform.
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    const Sheet& sheet = sheets[sheet_idx];
    for (const auto& [row, cells] : sheet.rows()) {
      for (std::size_t col = 0; col < cells.size(); ++col) {
        if (cells[col].formula_text.empty()) {
          continue;
        }
        const FormulaRewriteResult result = rewrite_formula(cells[col].formula_text, transform);
        if (result.changed) {
          cell_updates.push_back(CellUpdate{sheet_idx, row, static_cast<std::uint32_t>(col), result.text});
        }
      }
    }
  }
  for (CellUpdate& update : cell_updates) {
    sheets[update.sheet_index].set_cell_formula(update.row, update.col, std::move(update.formula));
    // A rename preserves workbook-relative sheet ids and therefore keeps the
    // existing dependency edges valid. Dirtying the changed owner is enough
    // to propagate through those edges; removal performs one final full
    // reindex after the sheet vector and metadata reach their final shape.
    mutator.mark_dirty(eval::CellNodeId{static_cast<std::uint16_t>(update.sheet_index), update.row, update.col});
  }

  for (DefinedName& entry : defined_names) {
    const FormulaRewriteResult result = rewrite_formula(entry.formula, transform);
    if (result.changed) {
      entry.formula = result.text;
      defined_names_changed = true;
    }
  }

  // A sheet mutation applies the same policy everywhere, so every slot of the
  // per-sheet table points at the single transform the caller supplied.
  const std::vector<const parser::RefTransform*> uniform(sheets.size(), &transform);
  rewrite_sheet_metadata_formulas(sheets, uniform, tables, pivot_caches, direct_sheet_old, direct_sheet_new,
                                  removed_sheet_name, dropped_cache_ids);
}

bool rewrite_defined_names(std::vector<DefinedName>& names, const parser::RefTransform& transform) {
  bool any_changed = false;
  for (DefinedName& entry : names) {
    FormulaRewriteResult result = rewrite_formula(entry.formula, transform);
    if (result.changed) {
      entry.formula = std::move(result.text);
      any_changed = true;
    }
  }
  return any_changed;
}

}  // namespace formulon
