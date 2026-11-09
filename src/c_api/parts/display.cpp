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

using formulon::c_api::parts::check_cell_coord;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::ensure_calculated;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::publish_scratch;
using formulon::c_api::parts::value_from_fm;

namespace {

static_assert(FM_DISPLAY_STATUS_OK == static_cast<int>(formulon::text_format::DisplayStatus::kOk),
              "FM_DISPLAY_STATUS_OK must match DisplayStatus::kOk");
static_assert(FM_DISPLAY_STATUS_OVERFLOW == static_cast<int>(formulon::text_format::DisplayStatus::kOverflow),
              "FM_DISPLAY_STATUS_OVERFLOW must match DisplayStatus::kOverflow");
static_assert(FM_DISPLAY_STATUS_INVALID_FORMAT ==
                  static_cast<int>(formulon::text_format::DisplayStatus::kInvalidFormat),
              "FM_DISPLAY_STATUS_INVALID_FORMAT must match DisplayStatus::kInvalidFormat");

// Publishes a rendering through the handle's read scratch.
void publish(const fm_workbook_t* wb, formulon::text_format::DisplayText display, const char** out_text,
             int32_t* out_status) {
  *out_text = publish_scratch(wb, std::move(display.text));
  *out_status = static_cast<int32_t>(display.status);
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_display_text(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                 uint32_t col, const char** out_text, int32_t* out_status) {
  clear_last_error();
  if (out_text == nullptr || out_status == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  if (auto rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  const formulon::Workbook& workbook = wb->workbook();
  publish(wb, formulon::text_format::format_cell_for_display(workbook, workbook.sheet(sheet_index), row, col), out_text,
          out_status);
  return 0;
}

extern "C" fm_status_t fm_workbook_format_value(const fm_workbook_t* wb, const fm_value_t* value,
                                                const char* number_format, const char** out_text, int32_t* out_status) {
  clear_last_error();
  if (wb == nullptr || value == nullptr || number_format == nullptr || out_text == nullptr || out_status == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({number_format}); rc != 0) {
    return rc;
  }
  formulon::Value v = formulon::Value::blank();
  if (auto rc = value_from_fm(*value, &v); rc != 0) {
    return rc;
  }
  publish(wb,
          formulon::text_format::format_value_for_display(v, number_format, wb->workbook().date1904(),
                                                          wb->workbook().excel_profile()),
          out_text, out_status);
  return 0;
}
