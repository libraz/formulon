// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the StructuredLog emitter. The output format is a single
// line of JSON written to stderr; we hand-roll the escaping to avoid pulling
// any external dependency into this low-level TU (see CLAUDE.md "Dependencies
// (strict)"). The emitter is lock-free: interleaved output from multiple
// threads may mix across records, but each record is emitted with a single
// `fwrite` call so lines remain intact on POSIX platforms.

#include "utils/structured_log.h"

#include <cmath>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>

#include "utils/error.h"

namespace formulon {
namespace {

/// Appends the JSON-escaped form of `in` (including surrounding quotes) to
/// `out`. Handles the standard JSON control-character set.
void AppendEscapedJsonString(std::string& out, std::string_view in) {
  out.push_back('"');
  for (char raw : in) {
    unsigned char c = static_cast<unsigned char>(raw);
    switch (c) {
      case '"':
        out.append("\\\"");
        break;
      case '\\':
        out.append("\\\\");
        break;
      case '\b':
        out.append("\\b");
        break;
      case '\f':
        out.append("\\f");
        break;
      case '\n':
        out.append("\\n");
        break;
      case '\r':
        out.append("\\r");
        break;
      case '\t':
        out.append("\\t");
        break;
      default:
        if (c < 0x20) {
          char buf[8];
          std::snprintf(buf, sizeof(buf), "\\u%04x", c);
          out.append(buf);
        } else {
          out.push_back(static_cast<char>(c));
        }
        break;
    }
  }
  out.push_back('"');
}

/// Appends an integer field separator and key; returns after writing the
/// leading `,"key":` portion.
void AppendKey(std::string& out, std::string_view key) {
  out.push_back(',');
  AppendEscapedJsonString(out, key);
  out.push_back(':');
}

}  // namespace

StructuredLog::StructuredLog(std::string_view event) : event_(event) {
  fields_.reserve(128);
}

StructuredLog& StructuredLog::field(std::string_view key, std::string_view value) {
  AppendKey(fields_, key);
  AppendEscapedJsonString(fields_, value);
  return *this;
}

StructuredLog& StructuredLog::field(std::string_view key, int64_t value) {
  AppendKey(fields_, key);
  char buf[32];
  std::snprintf(buf, sizeof(buf), "%lld", static_cast<long long>(value));
  fields_.append(buf);
  return *this;
}

StructuredLog& StructuredLog::field(std::string_view key, double value) {
  AppendKey(fields_, key);
  if (std::isnan(value)) {
    fields_.append("\"NaN\"");
  } else if (std::isinf(value)) {
    fields_.append(value < 0 ? "\"-Infinity\"" : "\"Infinity\"");
  } else {
    char buf[64];
    std::snprintf(buf, sizeof(buf), "%.17g", value);
    fields_.append(buf);
  }
  return *this;
}

StructuredLog& StructuredLog::field(std::string_view key, bool value) {
  AppendKey(fields_, key);
  fields_.append(value ? "true" : "false");
  return *this;
}

StructuredLog& StructuredLog::error_code(FormulonErrorCode code) {
  AppendKey(fields_, "code");
  char buf[32];
  std::snprintf(buf, sizeof(buf), "%d", static_cast<int>(code));
  fields_.append(buf);
  AppendKey(fields_, "code_name");
  AppendEscapedJsonString(fields_, to_cstring(code));
  return *this;
}

void StructuredLog::flush(std::string_view level) {
  std::string out;
  out.reserve(event_.size() + fields_.size() + 48);
  out.append("{\"level\":");
  AppendEscapedJsonString(out, level);
  out.append(",\"event\":");
  AppendEscapedJsonString(out, event_);
  out.append(fields_);
  out.push_back('}');
  out.push_back('\n');
  std::fwrite(out.data(), 1, out.size(), stderr);
  std::fflush(stderr);
  fields_.clear();
}

void StructuredLog::debug() {
  flush("debug");
}
void StructuredLog::info() {
  flush("info");
}
void StructuredLog::warn() {
  flush("warn");
}
void StructuredLog::error() {
  flush("error");
}

}  // namespace formulon
