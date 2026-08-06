//
// Implementation of `cli/render.h`. See header for contract.

#include "cli/render.h"

#include <cstdint>
#include <cstdio>
#include <string>

#include "c_api/formulon_c.h"
#include "utils/a1_column.h"
#include "utils/double_format.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace cli {

void append_column_letters(std::string& out, std::uint32_t col) {
  FM_CHECK(a1::append_column_letters(out, col), "column is outside Excel's grid");
}

std::string format_a1(std::uint32_t row, std::uint32_t col) {
  std::string out;
  append_column_letters(out, col);
  out.append(std::to_string(row + 1));
  return out;
}

namespace {

// Maps the int32 ordinal carried in `fm_value_t::u.error_code` back
// into the Excel-visible display string. Falls back to `#UNKNOWN!` for
// any out-of-range ordinal so the CLI never emits raw integers.
const char* error_display(std::int32_t ordinal) {
  if (ordinal < 0 || ordinal > static_cast<std::int32_t>(ErrorCode::Unknown)) {
    return "#UNKNOWN!";
  }
  return display_name(static_cast<ErrorCode>(ordinal));
}

// Appends a JSON-escaped form of `s` (without surrounding quotes) to
// `out`. Handles the standard escape set; non-ASCII bytes pass through
// unchanged because JSON is UTF-8 by default.
void append_json_escaped(std::string& out, std::string_view text) {
  for (char raw : text) {
    const unsigned char c = static_cast<unsigned char>(raw);
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
        if (c < 0x20U) {
          char esc[8];
          // Stable C escape for ASCII control bytes.
          std::snprintf(esc, sizeof(esc), "\\u%04x", c);
          out.append(esc);
        } else {
          out.push_back(static_cast<char>(c));
        }
        break;
    }
  }
}

}  // namespace

std::string render_value(const fm_value_t& v) {
  std::string out;
  switch (v.kind) {
    case FM_VAL_BLANK:
      return out;
    case FM_VAL_NUMBER:
      format_double(out, v.u.number);
      return out;
    case FM_VAL_BOOL:
      out.append(v.u.boolean != 0 ? "TRUE" : "FALSE");
      return out;
    case FM_VAL_TEXT:
      if (v.u.text != nullptr) {
        out.append(v.u.text);
      }
      return out;
    case FM_VAL_ERROR:
      out.append(error_display(v.u.error_code));
      return out;
    case FM_VAL_ARRAY:
      out.append("[array]");
      return out;
    case FM_VAL_REF:
      out.append("[ref]");
      return out;
    case FM_VAL_LAMBDA:
      out.append("[lambda]");
      return out;
  }
  return out;
}

std::string escape_single_line(std::string_view text) {
  std::string out;
  out.reserve(text.size());
  append_json_escaped(out, text);
  return out;
}

std::string render_value_json(const fm_value_t& v) {
  std::string out;
  out.push_back('{');
  out.append("\"kind\":\"");
  switch (v.kind) {
    case FM_VAL_BLANK:
      out.append("blank");
      break;
    case FM_VAL_NUMBER:
      out.append("number");
      break;
    case FM_VAL_BOOL:
      out.append("bool");
      break;
    case FM_VAL_TEXT:
      out.append("text");
      break;
    case FM_VAL_ERROR:
      out.append("error");
      break;
    case FM_VAL_ARRAY:
      out.append("array");
      break;
    case FM_VAL_REF:
      out.append("ref");
      break;
    case FM_VAL_LAMBDA:
      out.append("lambda");
      break;
  }
  out.append("\",\"value\":");
  switch (v.kind) {
    case FM_VAL_BLANK:
      out.append("null");
      break;
    case FM_VAL_NUMBER:
      format_double(out, v.u.number);
      break;
    case FM_VAL_BOOL:
      out.append(v.u.boolean != 0 ? "true" : "false");
      break;
    case FM_VAL_TEXT:
      out.push_back('"');
      append_json_escaped(out, v.u.text != nullptr ? std::string_view(v.u.text) : std::string_view{});
      out.push_back('"');
      break;
    case FM_VAL_ERROR:
      out.push_back('"');
      out.append(error_display(v.u.error_code));
      out.push_back('"');
      break;
    case FM_VAL_ARRAY:
    case FM_VAL_REF:
    case FM_VAL_LAMBDA:
      out.append("null");
      break;
  }
  out.push_back('}');
  return out;
}

}  // namespace cli
}  // namespace formulon
