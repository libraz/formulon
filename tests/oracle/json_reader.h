// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tiny, test-only JSON reader for Formulon's oracle golden files.
//
// The shape we need to parse is deliberately small: object, array, string,
// number (double), boolean, null. Golden JSON files are produced by the
// Python oracle-gen pipeline (`tools/oracle/oracle_gen.py`), so format
// assumptions — trailing commas forbidden, NaN / Infinity forbidden,
// duplicate keys undefined — follow Python's `json.dumps` output.
//
// This reader is intentionally not in `src/`: it is a test helper and must
// not leak into the engine (engine deps are restricted to the six libraries
// listed in CLAUDE.md, and no JSON library is one of them).

#ifndef FORMULON_TESTS_ORACLE_JSON_READER_H_
#define FORMULON_TESTS_ORACLE_JSON_READER_H_

#include <cstdint>
#include <map>
#include <memory>
#include <string>
#include <string_view>
#include <vector>

#include "utils/expected.h"

namespace formulon {
namespace tests {
namespace oracle {

class JsonValue;

/// One of the six JSON value kinds we care about. Arrays and objects own
/// their children; strings own their decoded buffer so string_view-style
/// aliasing into the source document is not required.
enum class JsonKind : std::uint8_t {
  Null,
  Bool,
  Number,
  String,
  Array,
  Object,
};

/// Minimal discriminated-union JSON value. Copyable; interior nodes are
/// heap-allocated via `unique_ptr` so self-references are impossible and
/// copy cost is O(size).
class JsonValue {
 public:
  JsonValue() : kind_(JsonKind::Null) {}

  static JsonValue make_null() { return JsonValue(); }
  static JsonValue make_bool(bool b);
  static JsonValue make_number(double n);
  static JsonValue make_string(std::string s);
  static JsonValue make_array(std::vector<JsonValue> items);
  static JsonValue make_object(std::map<std::string, JsonValue> members);

  JsonValue(const JsonValue& other) { copy_from(other); }
  JsonValue& operator=(const JsonValue& other) {
    if (this != &other) copy_from(other);
    return *this;
  }
  JsonValue(JsonValue&&) noexcept = default;
  JsonValue& operator=(JsonValue&&) noexcept = default;
  ~JsonValue() = default;

  JsonKind kind() const noexcept { return kind_; }
  bool is_null() const noexcept { return kind_ == JsonKind::Null; }
  bool is_bool() const noexcept { return kind_ == JsonKind::Bool; }
  bool is_number() const noexcept { return kind_ == JsonKind::Number; }
  bool is_string() const noexcept { return kind_ == JsonKind::String; }
  bool is_array() const noexcept { return kind_ == JsonKind::Array; }
  bool is_object() const noexcept { return kind_ == JsonKind::Object; }

  bool as_bool() const noexcept { return bool_; }
  double as_number() const noexcept { return number_; }
  const std::string& as_string() const noexcept { return *string_; }
  const std::vector<JsonValue>& as_array() const noexcept { return *array_; }
  const std::map<std::string, JsonValue>& as_object() const noexcept {
    return *object_;
  }

  /// Returns a pointer to the value at `key`, or `nullptr` if missing. Must
  /// be called only on an object.
  const JsonValue* find(std::string_view key) const;

 private:
  void copy_from(const JsonValue& other);

  JsonKind kind_;
  bool bool_ = false;
  double number_ = 0.0;
  std::unique_ptr<std::string> string_;
  std::unique_ptr<std::vector<JsonValue>> array_;
  std::unique_ptr<std::map<std::string, JsonValue>> object_;
};

/// Error reported by the reader. `line` / `column` are 1-based positions
/// into the source (useful in test failure messages).
struct JsonError {
  std::string message;
  std::size_t line = 1;
  std::size_t column = 1;
};

/// Parses `source` as a single JSON document. Trailing whitespace (including
/// a trailing newline, which `json.dumps` does not add but common editors
/// do) is tolerated. Any byte after the root value that is not whitespace
/// is an error. UTF-8 pass-through in strings is supported; `\u` escapes
/// are decoded to UTF-8 for the BMP and via surrogate pairs for SMP.
Expected<JsonValue, JsonError> parse_json(std::string_view source);

/// Convenience wrapper for test fixtures: reads `path` from disk and parses
/// it. Failures (file open, parse) are reported as a `JsonError` with
/// `message` prefixed by the failing stage.
Expected<JsonValue, JsonError> parse_json_file(const std::string& path);

}  // namespace oracle
}  // namespace tests
}  // namespace formulon

#endif  // FORMULON_TESTS_ORACLE_JSON_READER_H_
