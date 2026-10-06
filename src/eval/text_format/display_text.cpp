#include "eval/text_format/display_text.h"

#include <cmath>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/text_format/number_format.h"
#include "sheet.h"
#include "style_resolve.h"
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
  const std::string_view effective = code.empty() ? std::string_view("General") : code;
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
  const std::optional<std::string_view> code = effective_num_fmt(styles, xf->num_fmt_id);
  return code ? *code : std::string_view{};
}

DisplayText format_cell_for_display(const Workbook& workbook, const Sheet& sheet, std::uint32_t row,
                                    std::uint32_t col) {
  const EffectiveXf choice = select_effective_xf(sheet, row, col);
  const std::vector<CellXf>& xfs = workbook.styles().cell_xfs;
  const CellXf* xf = nullptr;
  if (!xfs.empty()) {
    const std::uint32_t xf_index = choice.xf_index < xfs.size() ? choice.xf_index : 0U;
    xf = &xfs[xf_index];
  }
  return format_value_for_display(sheet.resolve_cell_value(row, col), number_format_code_for_xf(workbook.styles(), xf),
                                  workbook.date1904());
}

}  // namespace text_format
}  // namespace formulon
