//
// Structured logging for Formulon. Every log record is a single line of JSON
// emitted on stderr by default, which keeps the dependency footprint zero (no
// spdlog, no fmt) while staying machine-parseable. Embedders can route records
// to their own sink or suppress levels they do not need.
//
// Typical usage:
//   StructuredLog("cell.evaluated")
//     .field("sheet", "Sheet1")
//     .field("a1", "B2")
//     .field("duration_us", 123)
//     .info();

#ifndef FORMULON_UTILS_STRUCTURED_LOG_H_
#define FORMULON_UTILS_STRUCTURED_LOG_H_

#include <cstdint>
#include <string>
#include <string_view>

#include "utils/error.h"

namespace formulon {

/// Severity threshold used by the process-wide structured-log configuration.
/// Records below the configured minimum are discarded before serialisation.
enum class StructuredLogLevel : uint8_t {
  kDebug,
  kInfo,
  kWarn,
  kError,
  kOff,
};

/// Receives one complete, newline-terminated JSON record synchronously.
/// The `record` view is valid only for the duration of this callback.
using StructuredLogSink = void (*)(std::string_view record, void* user_data);

/// Routes subsequent records to `sink`. Passing nullptr restores stderr.
/// The caller owns `user_data` and must keep it valid while the sink is set.
void set_structured_log_sink(StructuredLogSink sink, void* user_data = nullptr);

/// Sets the lowest severity emitted process-wide. The default is `kDebug`.
void set_structured_log_min_level(StructuredLogLevel level);

/// Fluent builder that emits a single structured log record.
///
/// The `StructuredLog` instance accumulates fields in an internal buffer and
/// flushes them when one of the terminal methods (`debug`, `info`, `warn`,
/// `error`) is called. If no terminal method is invoked the record is
/// discarded silently; callers must remember to complete the chain.
class StructuredLog {
 public:
  /// Begins a log record for the given event name.
  ///
  /// The event name should use a dotted namespace (e.g. `"parser.error"`).
  explicit StructuredLog(std::string_view event);

  // Non-copyable; logs are transient and always moved through the builder.
  StructuredLog(const StructuredLog&) = delete;
  StructuredLog& operator=(const StructuredLog&) = delete;
  StructuredLog(StructuredLog&&) = default;
  StructuredLog& operator=(StructuredLog&&) = default;
  ~StructuredLog() = default;

  /// Appends a string-valued field.
  StructuredLog& field(std::string_view key, std::string_view value);
  /// Appends a signed integer field.
  StructuredLog& field(std::string_view key, int64_t value);
  /// Appends a double-valued field. NaN and infinity are stringified.
  StructuredLog& field(std::string_view key, double value);
  /// Appends a boolean field.
  StructuredLog& field(std::string_view key, bool value);

  /// Records the `FormulonErrorCode` as both a numeric `"code"` field and
  /// a textual `"code_name"` field.
  StructuredLog& error_code(FormulonErrorCode code);

  /// Flushes the record at DEBUG severity.
  void debug();
  /// Flushes the record at INFO severity.
  void info();
  /// Flushes the record at WARN severity.
  void warn();
  /// Flushes the record at ERROR severity.
  void error();

 private:
  void flush(StructuredLogLevel level, std::string_view level_name);

  std::string event_;
  std::string fields_;  // Pre-serialised `,"k":v,"k2":v2` fragment.
};

}  // namespace formulon

#endif  // FORMULON_UTILS_STRUCTURED_LOG_H_
