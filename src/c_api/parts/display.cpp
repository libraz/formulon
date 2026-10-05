//
// C ABI - cell display text and ad-hoc number-format rendering.
//

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/text_format/display_text.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

using formulon::c_api::parts::check_finite;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::TextStore;

namespace {

static_assert(FM_DISPLAY_OK == static_cast<int>(formulon::text_format::DisplayStatus::kOk),
              "FM_DISPLAY_OK must match DisplayStatus::kOk");
static_assert(FM_DISPLAY_OVERFLOW == static_cast<int>(formulon::text_format::DisplayStatus::kOverflow),
              "FM_DISPLAY_OVERFLOW must match DisplayStatus::kOverflow");
static_assert(FM_DISPLAY_INVALID_FORMAT == static_cast<int>(formulon::text_format::DisplayStatus::kInvalidFormat),
              "FM_DISPLAY_INVALID_FORMAT must match DisplayStatus::kInvalidFormat");

// Publishes a rendering through the handle's read scratch.
void publish(const fm_workbook_t* wb, formulon::text_format::DisplayText display, const char** out_text,
             int32_t* out_status) {
  TextStore& store = const_cast<TextStore&>(wb->read_scratch);
  store.clear();
  store.emplace_back(std::move(display.text));
  *out_text = store.back().c_str();
  *out_status = static_cast<int32_t>(display.status);
}

// Converts a caller-supplied scalar into a `Value`. A Text value aliases the
// caller's string, which outlives the rendering call.
fm_status_t value_from_fm(const fm_value_t& in, formulon::Value* out, const char* api) {
  switch (in.kind) {
    case FM_VAL_BLANK:
      *out = formulon::Value::blank();
      return 0;
    case FM_VAL_NUMBER:
      if (auto rc = check_finite(in.u.number, api, "value"); rc != 0) {
        return rc;
      }
      *out = formulon::Value::number(in.u.number);
      return 0;
    case FM_VAL_BOOL:
      *out = formulon::Value::boolean(in.u.boolean != 0);
      return 0;
    case FM_VAL_TEXT:
      if (in.u.text == nullptr) {
        return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "value text is NULL", api);
      }
      *out = formulon::Value::text(in.u.text);
      return 0;
    case FM_VAL_ERROR:
      if (in.u.error_code < 0 || in.u.error_code > static_cast<int32_t>(formulon::ErrorCode::Unknown)) {
        return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "error code out of range",
                                 std::string(api) + ": error=" + std::to_string(in.u.error_code));
      }
      *out = formulon::Value::error(static_cast<formulon::ErrorCode>(in.u.error_code));
      return 0;
    default:
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "value kind not renderable",
                               std::string(api) + ": kind=" + std::to_string(static_cast<int>(in.kind)));
  }
}

}  // namespace

extern "C" fm_status_t fm_workbook_get_display_text(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                    uint32_t col, const char** out_text, int32_t* out_status) {
  clear_last_error();
  if (out_text == nullptr || out_status == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_get_display_text: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_get_display_text"); rc != 0) {
    return rc;
  }
  if (!formulon::Sheet::coord_in_grid(row, col)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_get_display_text: cell coordinate out of range",
                             "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  const formulon::Workbook& workbook = wb->workbook();
  publish(wb, formulon::text_format::format_cell_for_display(workbook, workbook.sheet(sheet_index), row, col), out_text,
          out_status);
  return 0;
}

extern "C" fm_status_t fm_workbook_format_value(const fm_workbook_t* wb, const fm_value_t* value,
                                                const char* format_code, const char** out_text, int32_t* out_status) {
  clear_last_error();
  if (wb == nullptr || value == nullptr || format_code == nullptr || out_text == nullptr || out_status == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_format_value: NULL argument");
  }
  formulon::Value v = formulon::Value::blank();
  if (auto rc = value_from_fm(*value, &v, "fm_workbook_format_value"); rc != 0) {
    return rc;
  }
  publish(wb, formulon::text_format::format_value_for_display(v, format_code, wb->workbook().date1904()), out_text,
          out_status);
  return 0;
}
