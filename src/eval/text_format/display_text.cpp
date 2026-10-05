#include "eval/text_format/display_text.h"

#include <cmath>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/text_format/number_format.h"
#include "eval/text_format/number_format_types.h"
#include "sheet.h"
#include "styles.h"
#include "workbook.h"

namespace formulon {
namespace text_format {
namespace {

constexpr std::string_view kOverflowText = "########";

DisplayText render_text(std::string_view text, std::string_view code) {
  DisplayText result;
  if (apply_text_format(text, code, result.text, FormatDialect::kStored) != FormatStatus::kOk) {
    result.text = std::string(text);
    result.status = DisplayStatus::kInvalidFormat;
  }
  return result;
}

DisplayText render_number(double number, std::string_view code, bool date1904) {
  DisplayText result;
  if (!std::isfinite(number)) {
    result.text = std::string(kOverflowText);
    result.status = DisplayStatus::kOverflow;
    return result;
  }
  // A text-only format (`@`) shows numbers as General.
  const auto sections = number_format_detail::split_sections(code);
  bool general_only = code.empty();
  if (sections.size() == 1U) {
    number_format_detail::Section section;
    number_format_detail::tokenize_section(sections[0], section, FormatDialect::kStored);
    number_format_detail::classify(section, sections[0]);
    general_only = general_only || section.is_text;
  }
  const std::string_view effective = general_only ? std::string_view("General") : code;
  std::string out;
  switch (apply_format(number, effective, out, date1904, FormatDialect::kStored)) {
    case FormatStatus::kOk:
      result.text = std::move(out);
      break;
    case FormatStatus::kOverflow:
      result.text = std::string(kOverflowText);
      result.status = DisplayStatus::kOverflow;
      break;
    case FormatStatus::kValueError:
      apply_format(number, "General", result.text, date1904, FormatDialect::kStored);
      result.status = DisplayStatus::kInvalidFormat;
      break;
  }
  return result;
}

}  // namespace

DisplayText format_value_for_display(const Value& value, std::string_view code, bool date1904) {
  switch (value.kind()) {
    case ValueKind::Number:
      return render_number(value.as_number(), code, date1904);
    case ValueKind::Text:
      return render_text(value.as_text(), code);
    case ValueKind::Bool:
      return render_text(value.as_boolean() ? "TRUE" : "FALSE", code);
    case ValueKind::Error: {
      DisplayText result;
      result.text = display_name(value.as_error());
      return result;
    }
    default:
      return DisplayText{};
  }
}

std::string_view number_format_code_for_xf(const StylesTable& styles, const CellXf* xf) {
  if (xf == nullptr) {
    return {};
  }
  for (const NumFmtRecord& custom : styles.num_fmts) {
    if (custom.id == xf->num_fmt_id && custom.format_string_index < styles.num_fmt_strings.size()) {
      return styles.num_fmt_strings[custom.format_string_index];
    }
  }
  const char* builtin = builtin_num_fmt(xf->num_fmt_id);
  return builtin == nullptr ? std::string_view{} : std::string_view(builtin);
}

DisplayText format_cell_for_display(const Workbook& workbook, const Sheet& sheet, std::uint32_t row,
                                    std::uint32_t col) {
  const Cell* cell = sheet.cell_at(row, col);
  if (cell == nullptr) {
    return DisplayText{};
  }
  const std::vector<CellXf>& xfs = workbook.styles().cell_xfs;
  const CellXf* xf = cell->xf_index < xfs.size() ? &xfs[cell->xf_index] : nullptr;
  return format_value_for_display(cell->cached_value, number_format_code_for_xf(workbook.styles(), xf),
                                  workbook.date1904());
}

}  // namespace text_format
}  // namespace formulon
