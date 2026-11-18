#include "io/package_reader_common.h"

#include <cmath>
#include <optional>
#include <string>
#include <unordered_set>
#include <utility>

#include "column_width.h"
#include "external_link.h"
#include "font_completion.h"
#include "io/dynamic_array_formula.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml_defs.h"
#include "io/sheet_phonetic.h"
#include "io/theme_part.h"
#include "sheet_layout_edit.h"

namespace formulon {
namespace io {

void mark_sheet_dynamic_arrays(
    Sheet& sheet, std::uint32_t dynamic_array_entry,
    const std::vector<std::tuple<std::uint32_t, std::uint32_t, std::uint32_t>>& cell_metadata) {
  std::unordered_set<std::uint64_t> marked;
  for (const auto& [row, col, entry] : cell_metadata) {
    if (dynamic_array_entry != 0U && entry == dynamic_array_entry) {
      marked.insert(dynamic_array_cell_key(row, col));
    }
  }
  apply_loaded_dynamic_array_marks(sheet, marked);
}

void mark_model_derived_parts(const Workbook& wb, std::vector<PassthroughPart>& parts) {
  std::unordered_set<std::string> derived;
  for (const ExternalLinkRecord& link : wb.external_links()) {
    if (link.kind == ExternalLinkRecord::Kind::kExternalBook && !link.part_path.empty()) {
      derived.insert(link.part_path);
      derived.insert(ooxml::rels_path_for_part(link.part_path));
    }
  }
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    const Sheet& sheet = wb.sheet(s);
    if (!sheet.comment_vml_path().empty()) {
      derived.insert(sheet.comment_vml_path());
    }
    for (const UnknownRelationship& rel : sheet.unknown_relationships()) {
      if (rel.type == kRelBinaryIndex && !rel.target_external) {
        derived.insert(rel.target);
      }
    }
  }
  for (PassthroughPart& part : parts) {
    if (derived.count(part.path) != 0U) {
      part.retained_origin = PassthroughPart::RetainedOrigin::kDerivedFromModel;
    }
  }
}

Error sheet_add_error(Error error, std::string_view message, std::string context) {
  if (error.code != FormulonErrorCode::kInvalidSheetName && error.code != FormulonErrorCode::kAlreadyExists) {
    return error;
  }
  return make_error(FormulonErrorCode::kIoSheetCorrupt, std::string(message), std::move(context));
}

void complete_loaded_fonts(StylesTable& styles, const Workbook& wb, const ReadOptions& options) {
  const ExcelLocale locale = options.profile.has_value() ? options.profile->locale : wb.excel_profile().locale;
  for (FontRecord& font : styles.fonts) {
    complete_font_from_face(font, locale);
  }
}

namespace {

/// Clears the width of every column whose non-custom width is the sheet default, the width Excel writes on a
/// column that has none of its own. Compared in the 1/256 steps `.xlsb` stores.
void clear_default_column_widths(Workbook& wb) {
  std::optional<ColumnWidthModel> model;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    Sheet& sheet = wb.sheet(i);
    SheetLayout& layout = sheet.mutable_layout();
    bool cleared = false;
    for (ColumnLayout& column : layout.columns) {
      if (!column.has_width || column.custom_width) {
        continue;
      }
      if (!model) {
        model = resolve_column_width_model(wb, wb.excel_profile(), GeometryMode::kDisplay);
      }
      if (std::floor(column.width * 256.0) == std::floor(default_column_width_chars(sheet, *model) * 256.0)) {
        column.has_width = false;
        column.width = 0.0;
        cleared = true;
      }
    }
    if (cleared) {
      drop_stateless_columns(layout);
    }
  }
}

}  // namespace

void finish_loaded_workbook(Workbook& wb, const ReadOptions& options,
                            const std::vector<std::vector<CellAddress>>& recalc_marked) {
  wb.apply_legacy_implicit_intersections();
  clamp_sheet_phonetic_fonts(wb);
  wb.request_pivot_refresh_on_load();
  if (options.profile.has_value()) {
    wb.assign_loaded_profile(*options.profile);
  }
  resolve_scheme_font_names(wb);
  clear_default_column_widths(wb);
  wb.apply_loaded_recalc_marks(recalc_marked);
}

}  // namespace io
}  // namespace formulon
