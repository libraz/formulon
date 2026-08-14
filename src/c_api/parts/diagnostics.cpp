//
// C ABI - diagnostics (`fm_last_error_*`, `fm_status_string`, Excel error
// display names) and the process-wide structured-log configuration.
//
// Thin pass-through to the thread-local diagnostic buffers populated by
// `parts/common.cpp`. The error-code-to-text mapping leverages the
// existing `formulon::to_cstring(FormulonErrorCode)` helper.

#include <cstddef>
#include <mutex>
#include <string>
#include <string_view>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "utils/error.h"
#include "utils/structured_log.h"
#include "value.h"

extern "C" const char* fm_last_error_message(void) {
  return formulon::c_api::parts::last_error_message();
}

extern "C" const char* fm_last_error_context(void) {
  return formulon::c_api::parts::last_error_context();
}

extern "C" const char* fm_status_string(fm_status_t status) {
  return formulon::to_cstring(static_cast<formulon::FormulonErrorCode>(status));
}

extern "C" const char* fm_error_display_name(fm_error_code_t error) {
  constexpr auto kLastError = static_cast<fm_error_code_t>(formulon::ErrorCode::Unknown);
  if (error < 0 || error > kLastError) {
    return formulon::display_name(formulon::ErrorCode::Unknown);
  }
  return formulon::display_name(static_cast<formulon::ErrorCode>(error));
}

namespace {

// The engine's sink takes a `std::string_view`, which cannot cross the C
// ABI, so the C sink is kept here and reached through a fixed adapter. The
// pair is copied out under the lock and the callback runs after the lock is
// released, matching the engine's own rule that a sink may reconfigure
// logging from inside the call.
struct CLogSink {
  fm_log_sink_cb cb = nullptr;
  void* user_data = nullptr;
};

std::mutex& c_log_sink_mutex() {
  static std::mutex mutex;
  return mutex;
}

CLogSink& c_log_sink() {
  static CLogSink sink;
  return sink;
}

void c_log_sink_adapter(std::string_view record, void* /*user_data*/) {
  CLogSink snapshot;
  {
    const std::lock_guard<std::mutex> lock(c_log_sink_mutex());
    snapshot = c_log_sink();
  }
  if (snapshot.cb != nullptr) {
    snapshot.cb(record.data(), record.size(), snapshot.user_data);
  }
}

}  // namespace

extern "C" fm_status_t fm_set_log_min_level(int32_t level) {
  formulon::c_api::parts::clear_last_error();
  if (level < static_cast<int32_t>(FM_LOG_LEVEL_DEBUG) || level > static_cast<int32_t>(FM_LOG_LEVEL_OFF)) {
    return formulon::c_api::parts::set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                                                     "fm_set_log_min_level: unknown level",
                                                     "level=" + std::to_string(level));
  }
  formulon::set_structured_log_min_level(static_cast<formulon::StructuredLogLevel>(level));
  return 0;
}

extern "C" fm_status_t fm_set_log_sink(fm_log_sink_cb sink, void* user_data) {
  formulon::c_api::parts::clear_last_error();
  {
    const std::lock_guard<std::mutex> lock(c_log_sink_mutex());
    c_log_sink().cb = sink;
    c_log_sink().user_data = sink == nullptr ? nullptr : user_data;
  }
  // Install the adapter only while a C sink is registered, so clearing it
  // restores the engine's own stderr fallback rather than a shim that drops
  // every record.
  formulon::set_structured_log_sink(sink == nullptr ? nullptr : &c_log_sink_adapter, nullptr);
  return 0;
}
