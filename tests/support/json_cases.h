#pragma once

// Minimal JSON reader for the case tables under tools/api/*.json that the C,
// C++ and CLI suites share with the JS and Python drivers. Strict enough for
// those files only: well-formed input is assumed, and numbers go through
// strtod. String escapes, \uXXXX surrogate pairs included, decode to UTF-8.

#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <fstream>
#include <sstream>
#include <string>
#include <vector>

namespace formulon::test_support {

struct Json {
  enum class Kind { kNull, kBool, kNumber, kString, kArray, kObject };
  Kind kind = Kind::kNull;
  bool boolean = false;
  double number = 0.0;
  std::string text;
  std::vector<Json> items;
  std::vector<std::string> keys;

  const Json* Find(const std::string& key) const {
    for (std::size_t i = 0; i < keys.size(); ++i) {
      if (keys[i] == key) {
        return &items[i];
      }
    }
    return nullptr;
  }
  const Json& At(const std::string& key) const {
    static const Json kMissing;
    const Json* found = Find(key);
    return found != nullptr ? *found : kMissing;
  }
};

class JsonReader {
 public:
  explicit JsonReader(const std::string& src) : src_(src) {}
  Json Parse() { return Value(); }

 private:
  void Skip() {
    while (pos_ < src_.size() &&
           (src_[pos_] == ' ' || src_[pos_] == '\n' || src_[pos_] == '\r' || src_[pos_] == '\t')) {
      ++pos_;
    }
  }
  std::uint32_t Hex4() {
    std::uint32_t v = 0;
    for (int i = 0; i < 4 && pos_ < src_.size(); ++i, ++pos_) {
      const char h = src_[pos_];
      v <<= 4;
      if (h >= '0' && h <= '9') {
        v |= static_cast<std::uint32_t>(h - '0');
      } else if (h >= 'a' && h <= 'f') {
        v |= static_cast<std::uint32_t>(h - 'a' + 10);
      } else if (h >= 'A' && h <= 'F') {
        v |= static_cast<std::uint32_t>(h - 'A' + 10);
      }
    }
    return v;
  }
  static void AppendUtf8(std::string* out, std::uint32_t cp) {
    if (cp < 0x80) {
      out->push_back(static_cast<char>(cp));
    } else if (cp < 0x800) {
      out->push_back(static_cast<char>(0xC0 | (cp >> 6)));
      out->push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    } else if (cp < 0x10000) {
      out->push_back(static_cast<char>(0xE0 | (cp >> 12)));
      out->push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
      out->push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    } else {
      out->push_back(static_cast<char>(0xF0 | (cp >> 18)));
      out->push_back(static_cast<char>(0x80 | ((cp >> 12) & 0x3F)));
      out->push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
      out->push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    }
  }
  std::string String() {
    std::string out;
    ++pos_;  // opening quote
    while (pos_ < src_.size() && src_[pos_] != '"') {
      if (src_[pos_] != '\\') {
        out.push_back(src_[pos_++]);
        continue;
      }
      ++pos_;
      const char e = src_[pos_++];
      switch (e) {
        case 'n':
          out.push_back('\n');
          break;
        case 't':
          out.push_back('\t');
          break;
        case 'r':
          out.push_back('\r');
          break;
        case 'b':
          out.push_back('\b');
          break;
        case 'f':
          out.push_back('\f');
          break;
        case 'u': {
          std::uint32_t cp = Hex4();
          // A high surrogate followed by an escaped low surrogate is one code point.
          if (cp >= 0xD800 && cp <= 0xDBFF && src_.compare(pos_, 2, "\\u") == 0) {
            const std::size_t save = pos_;
            pos_ += 2;
            const std::uint32_t lo = Hex4();
            if (lo >= 0xDC00 && lo <= 0xDFFF) {
              cp = 0x10000 + ((cp - 0xD800) << 10) + (lo - 0xDC00);
            } else {
              pos_ = save;
            }
          }
          AppendUtf8(&out, cp);
          break;
        }
        default:
          out.push_back(e);
          break;
      }
    }
    ++pos_;
    return out;
  }
  Json Value() {
    Skip();
    Json v;
    const char c = src_[pos_];
    if (c == '{' || c == '[') {
      const char close = c == '{' ? '}' : ']';
      v.kind = c == '{' ? Json::Kind::kObject : Json::Kind::kArray;
      ++pos_;
      Skip();
      while (src_[pos_] != close) {
        if (c == '{') {
          Skip();
          v.keys.push_back(String());
          Skip();
          ++pos_;  // colon
        }
        v.items.push_back(Value());
        Skip();
        if (src_[pos_] == ',') {
          ++pos_;
          Skip();
        }
      }
      ++pos_;
    } else if (c == '"') {
      v.kind = Json::Kind::kString;
      v.text = String();
    } else if (src_.compare(pos_, 4, "true") == 0 || src_.compare(pos_, 5, "false") == 0) {
      v.kind = Json::Kind::kBool;
      v.boolean = c == 't';
      pos_ += v.boolean ? 4 : 5;
    } else if (src_.compare(pos_, 4, "null") == 0) {
      pos_ += 4;
    } else {
      v.kind = Json::Kind::kNumber;
      char* end = nullptr;
      v.number = std::strtod(src_.c_str() + pos_, &end);
      pos_ = static_cast<std::size_t>(end - src_.c_str());
    }
    return v;
  }

  const std::string& src_;
  std::size_t pos_ = 0;
};

// Loads `relative` (e.g. "tools/api/date_serial_cases.json") from the
// repository root, located from `this_file` (pass __FILE__) and the test
// directory it sits in (e.g. "tests/c_api/").
inline bool LoadRepoJson(const std::string& this_file, const std::string& test_dir, const std::string& relative,
                         Json* out) {
  const std::size_t at = this_file.rfind(test_dir);
  if (at == std::string::npos) {
    return false;
  }
  std::ifstream in(this_file.substr(0, at) + relative);
  if (!in) {
    return false;
  }
  std::stringstream buffer;
  buffer << in.rdbuf();
  const std::string src = buffer.str();
  *out = JsonReader(src).Parse();
  return true;
}

}  // namespace formulon::test_support
