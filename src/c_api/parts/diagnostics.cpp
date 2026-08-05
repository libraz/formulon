//
// C ABI - diagnostics (`fm_last_error_*`, `fm_status_string`).
//
// Thin pass-through to the thread-local diagnostic buffers populated by
// `parts/common.cpp`. The error-code-to-text mapping leverages the
// existing `formulon::to_cstring(FormulonErrorCode)` helper.

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "utils/error.h"

extern "C" const char* fm_last_error_message(void) {
  return formulon::c_api::parts::last_error_message();
}

extern "C" const char* fm_last_error_context(void) {
  return formulon::c_api::parts::last_error_context();
}

extern "C" const char* fm_status_string(fm_status_t status) {
  return formulon::to_cstring(static_cast<formulon::FormulonErrorCode>(status));
}
