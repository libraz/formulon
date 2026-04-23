// Copyright 2026 libraz. Licensed under the MIT License.

#include "tests/oracle/json_reader.h"

#include <cerrno>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <string>
#include <string_view>
#include <utility>

namespace formulon {
namespace tests {
namespace oracle {

// ---------------------------------------------------------------------------
// JsonValue factories / copy helper
// ---------------------------------------------------------------------------

JsonValue JsonValue::make_bool(bool b) {
  JsonValue v;
  v.kind_ = JsonKind::Bool;
  v.bool_ = b;
  return v;
}

JsonValue JsonValue::make_number(double n) {
  JsonValue v;
  v.kind_ = JsonKind::Number;
  v.number_ = n;
  return v;
}

JsonValue JsonValue::make_string(std::string s) {
  JsonValue v;
  v.kind_ = JsonKind::String;
  v.string_ = std::make_unique<std::string>(std::move(s));
  return v;
}

JsonValue JsonValue::make_array(std::vector<JsonValue> items) {
  JsonValue v;
  v.kind_ = JsonKind::Array;
  v.array_ = std::make_unique<std::vector<JsonValue>>(std::move(items));
  return v;
}

JsonValue JsonValue::make_object(std::map<std::string, JsonValue> members) {
  JsonValue v;
  v.kind_ = JsonKind::Object;
  v.object_ = std::make_unique<std::map<std::string, JsonValue>>(std::move(members));
  return v;
}

void JsonValue::copy_from(const JsonValue& other) {
  kind_ = other.kind_;
  bool_ = other.bool_;
  number_ = other.number_;
  string_ = other.string_ ? std::make_unique<std::string>(*other.string_) : nullptr;
  array_ = other.array_ ? std::make_unique<std::vector<JsonValue>>(*other.array_) : nullptr;
  object_ =
      other.object_ ? std::make_unique<std::map<std::string, JsonValue>>(*other.object_) : nullptr;
}

const JsonValue* JsonValue::find(std::string_view key) const {
  if (kind_ != JsonKind::Object || !object_) return nullptr;
  auto it = object_->find(std::string(key));
  return it == object_->end() ? nullptr : &it->second;
}

// ---------------------------------------------------------------------------
// Parser
// ---------------------------------------------------------------------------

namespace {

// Small scan state. Methods return Expected<T, JsonError>; errors carry
// (message, line, column) fixed at the failure point. Callers can return
// any `err(...)` directly thanks to the implicit E → Expected conversion.
class Parser {
 public:
  explicit Parser(std::string_view src) : src_(src) {}

  Expected<JsonValue, JsonError> parse_document() {
    skip_ws();
    auto root = parse_value();
    if (!root.has_value()) return root;
    skip_ws();
    if (pos_ != src_.size()) return err("trailing content after root value");
    return root;
  }

 private:
  Expected<JsonValue, JsonError> parse_value() {
    skip_ws();
    if (pos_ >= src_.size()) return err("unexpected end of input");
    char c = src_[pos_];
    switch (c) {
      case '{':
        return parse_object();
      case '[':
        return parse_array();
      case '"':
        return parse_string_value();
      case 't':
      case 'f':
        return parse_bool();
      case 'n':
        return parse_null();
      default:
        if (c == '-' || (c >= '0' && c <= '9')) return parse_number();
        return err("unexpected character");
    }
  }

  Expected<JsonValue, JsonError> parse_object() {
    advance();  // '{'
    std::map<std::string, JsonValue> members;
    skip_ws();
    if (peek() == '}') {
      advance();
      return JsonValue::make_object(std::move(members));
    }
    while (true) {
      skip_ws();
      if (peek() != '"') return err("expected string key");
      auto key = parse_string_raw();
      if (!key.has_value()) return key.error();
      skip_ws();
      if (peek() != ':') return err("expected ':' after key");
      advance();
      auto val = parse_value();
      if (!val.has_value()) return val;
      members.emplace(std::move(key.value()), std::move(val.value()));
      skip_ws();
      char sep = peek();
      if (sep == ',') {
        advance();
        continue;
      }
      if (sep == '}') {
        advance();
        return JsonValue::make_object(std::move(members));
      }
      return err("expected ',' or '}' in object");
    }
  }

  Expected<JsonValue, JsonError> parse_array() {
    advance();  // '['
    std::vector<JsonValue> items;
    skip_ws();
    if (peek() == ']') {
      advance();
      return JsonValue::make_array(std::move(items));
    }
    while (true) {
      auto val = parse_value();
      if (!val.has_value()) return val;
      items.push_back(std::move(val.value()));
      skip_ws();
      char sep = peek();
      if (sep == ',') {
        advance();
        continue;
      }
      if (sep == ']') {
        advance();
        return JsonValue::make_array(std::move(items));
      }
      return err("expected ',' or ']' in array");
    }
  }

  Expected<JsonValue, JsonError> parse_string_value() {
    auto s = parse_string_raw();
    if (!s.has_value()) return s.error();
    return JsonValue::make_string(std::move(s.value()));
  }

  Expected<std::string, JsonError> parse_string_raw() {
    if (peek() != '"') return err("expected '\"'");
    advance();
    std::string out;
    while (pos_ < src_.size()) {
      char c = src_[pos_];
      if (c == '"') {
        advance();
        return out;
      }
      if (c == '\\') {
        advance();
        if (pos_ >= src_.size()) return err("unterminated string escape");
        char esc = src_[pos_];
        advance();
        switch (esc) {
          case '"':
            out.push_back('"');
            break;
          case '\\':
            out.push_back('\\');
            break;
          case '/':
            out.push_back('/');
            break;
          case 'b':
            out.push_back('\b');
            break;
          case 'f':
            out.push_back('\f');
            break;
          case 'n':
            out.push_back('\n');
            break;
          case 'r':
            out.push_back('\r');
            break;
          case 't':
            out.push_back('\t');
            break;
          case 'u': {
            auto cp = parse_hex_codepoint();
            if (!cp.has_value()) return cp.error();
            std::uint32_t code = cp.value();
            if (code >= 0xD800 && code <= 0xDBFF) {
              // High surrogate -> consume matching low surrogate and combine
              // into an SMP code point.
              if (pos_ + 1 >= src_.size() || src_[pos_] != '\\' ||
                  src_[pos_ + 1] != 'u') {
                return err("expected low surrogate");
              }
              advance();
              advance();
              auto lo = parse_hex_codepoint();
              if (!lo.has_value()) return lo.error();
              std::uint32_t low = lo.value();
              if (low < 0xDC00 || low > 0xDFFF) return err("invalid low surrogate");
              code = 0x10000 + ((code - 0xD800) << 10) + (low - 0xDC00);
            }
            append_utf8(out, code);
            break;
          }
          default:
            return err("unknown string escape");
        }
      } else {
        out.push_back(c);
        advance();
      }
    }
    return err("unterminated string");
  }

  Expected<std::uint32_t, JsonError> parse_hex_codepoint() {
    if (pos_ + 4 > src_.size()) return err("short \\u escape");
    std::uint32_t v = 0;
    for (int i = 0; i < 4; ++i) {
      char c = src_[pos_ + static_cast<std::size_t>(i)];
      std::uint32_t d = 0;
      if (c >= '0' && c <= '9') {
        d = static_cast<std::uint32_t>(c - '0');
      } else if (c >= 'a' && c <= 'f') {
        d = 10U + static_cast<std::uint32_t>(c - 'a');
      } else if (c >= 'A' && c <= 'F') {
        d = 10U + static_cast<std::uint32_t>(c - 'A');
      } else {
        return err("invalid hex digit in \\u escape");
      }
      v = (v << 4) | d;
    }
    pos_ += 4;
    column_ += 4;
    return v;
  }

  static void append_utf8(std::string& out, std::uint32_t cp) {
    if (cp < 0x80) {
      out.push_back(static_cast<char>(cp));
    } else if (cp < 0x800) {
      out.push_back(static_cast<char>(0xC0 | (cp >> 6)));
      out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    } else if (cp < 0x10000) {
      out.push_back(static_cast<char>(0xE0 | (cp >> 12)));
      out.push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
      out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    } else {
      out.push_back(static_cast<char>(0xF0 | (cp >> 18)));
      out.push_back(static_cast<char>(0x80 | ((cp >> 12) & 0x3F)));
      out.push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
      out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    }
  }

  Expected<JsonValue, JsonError> parse_bool() {
    if (src_.substr(pos_, 4) == "true") {
      pos_ += 4;
      column_ += 4;
      return JsonValue::make_bool(true);
    }
    if (src_.substr(pos_, 5) == "false") {
      pos_ += 5;
      column_ += 5;
      return JsonValue::make_bool(false);
    }
    return err("invalid bool literal");
  }

  Expected<JsonValue, JsonError> parse_null() {
    if (src_.substr(pos_, 4) == "null") {
      pos_ += 4;
      column_ += 4;
      return JsonValue::make_null();
    }
    return err("invalid null literal");
  }

  Expected<JsonValue, JsonError> parse_number() {
    std::size_t start = pos_;
    if (peek() == '-') advance();
    while (pos_ < src_.size()) {
      char c = src_[pos_];
      if ((c >= '0' && c <= '9') || c == '.' || c == 'e' || c == 'E' || c == '+' ||
          c == '-') {
        advance();
      } else {
        break;
      }
    }
    if (pos_ == start) return err("empty number");
    std::string num{src_.substr(start, pos_ - start)};
    char* end = nullptr;
    errno = 0;
    double v = std::strtod(num.c_str(), &end);
    if (end != num.c_str() + num.size() || errno == ERANGE) {
      return err("malformed number");
    }
    return JsonValue::make_number(v);
  }

  void skip_ws() {
    while (pos_ < src_.size()) {
      char c = src_[pos_];
      if (c == ' ' || c == '\t' || c == '\r') {
        advance();
      } else if (c == '\n') {
        advance_newline();
      } else {
        return;
      }
    }
  }

  char peek() const { return pos_ < src_.size() ? src_[pos_] : '\0'; }

  void advance() {
    if (pos_ < src_.size()) {
      ++pos_;
      ++column_;
    }
  }

  void advance_newline() {
    if (pos_ < src_.size()) {
      ++pos_;
      ++line_;
      column_ = 1;
    }
  }

  // Note: returns a plain JsonError (not an Expected). Thanks to
  // `Expected<T, E>::Expected(E)` being non-explicit, callers can
  // `return err("...")` from any `Expected<T, JsonError>` function.
  JsonError err(const char* message) const {
    JsonError e;
    e.message = message;
    e.line = line_;
    e.column = column_;
    return e;
  }

  std::string_view src_;
  std::size_t pos_ = 0;
  std::size_t line_ = 1;
  std::size_t column_ = 1;
};

}  // namespace

Expected<JsonValue, JsonError> parse_json(std::string_view source) {
  Parser p(source);
  return p.parse_document();
}

Expected<JsonValue, JsonError> parse_json_file(const std::string& path) {
  FILE* fp = std::fopen(path.c_str(), "rb");
  if (fp == nullptr) {
    JsonError e;
    e.message = std::string("cannot open: ") + path;
    return e;
  }
  std::string buffer;
  char chunk[4096];
  while (true) {
    std::size_t n = std::fread(chunk, 1, sizeof(chunk), fp);
    if (n == 0) break;
    buffer.append(chunk, n);
  }
  std::fclose(fp);
  return parse_json(buffer);
}

}  // namespace oracle
}  // namespace tests
}  // namespace formulon
