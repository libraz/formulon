//
// Test-side capture of structured-log records.
//
// Formulon ships with logging off (`StructuredLogLevel::kOff`) so an embedded
// library never writes to its host's stderr unasked. That makes stderr-capture
// assertions silently vacuous: with the threshold at `kOff` the level gate
// rejects every record before it reaches any sink, so "stderr does not contain
// X" holds no matter what the code does — including when the whole logging
// path is broken.
//
// `LogRecorder` closes that hole. It opts logging in the way an embedder does,
// installs a sink, and keeps the records it receives, so a test asserts
// against the record list rather than against a stream that may never have
// been written. Absence assertions become "the sink received nothing", which
// fails both when an unwanted record IS emitted and when the mechanism that
// would carry it has stopped working.
//
// Because absence alone can still pass for the wrong reason, `probe_and_clear`
// emits one record through the production builder and reports whether it
// arrived. A test that calls it before the code under test has proven the sink
// is live at that moment.
//
// The structured-log configuration is process-wide, so the recorder restores
// the shipped defaults (no sink, `kOff`) in its destructor. Construct one per
// test; do not hold two at once.

#ifndef FORMULON_TESTS_UTIL_TEST_LOG_RECORDER_H_
#define FORMULON_TESTS_UTIL_TEST_LOG_RECORDER_H_

#include <string>
#include <string_view>
#include <vector>

#include "utils/structured_log.h"

namespace formulon {
namespace test {

class LogRecorder {
 public:
  /// Installs the sink and lowers the process-wide threshold to `level`
  /// (`kDebug` by default, i.e. everything is captured).
  explicit LogRecorder(StructuredLogLevel level = StructuredLogLevel::kDebug) {
    set_structured_log_sink(&LogRecorder::Append, this);
    set_structured_log_min_level(level);
  }

  /// Restores the shipped defaults: threshold `kOff` and no sink.
  ~LogRecorder() {
    set_structured_log_min_level(StructuredLogLevel::kOff);
    set_structured_log_sink(nullptr);
  }

  LogRecorder(const LogRecorder&) = delete;
  LogRecorder& operator=(const LogRecorder&) = delete;

  /// Emits one record through the production builder and reports whether the
  /// sink received it, then drops every captured record. A `false` return
  /// means the logging path is not carrying records at all, which is exactly
  /// the state that makes a bare absence assertion meaningless.
  bool probe_and_clear() {
    const std::size_t before = records_.size();
    StructuredLog("test.log_recorder.probe").debug();
    const bool delivered = records_.size() > before;
    clear();
    return delivered;
  }

  /// Drops every captured record.
  void clear() { records_.clear(); }

  /// Records captured since construction or the last `clear()`, in order.
  const std::vector<std::string>& records() const { return records_; }

  /// True when no record has been captured.
  bool empty() const { return records_.empty(); }

  /// True when some captured record contains `needle`.
  bool contains(std::string_view needle) const {
    for (const std::string& record : records_) {
      if (record.find(needle) != std::string::npos) {
        return true;
      }
    }
    return false;
  }

  /// The captured records joined for a failure message; empty when none were
  /// captured.
  std::string joined() const {
    std::string out;
    for (const std::string& record : records_) {
      out += record;
    }
    return out;
  }

 private:
  static void Append(std::string_view record, void* user_data) {
    static_cast<LogRecorder*>(user_data)->records_.emplace_back(record);
  }

  std::vector<std::string> records_;
};

}  // namespace test
}  // namespace formulon

#endif  // FORMULON_TESTS_UTIL_TEST_LOG_RECORDER_H_
