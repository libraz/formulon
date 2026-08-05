
#include "io/xml_escape.h"

#include <cstdint>
#include <string>
#include <string_view>

namespace formulon {
namespace io {
namespace {

bool IsOoxmlEscape(std::string_view in, std::size_t pos, std::uint16_t* code) {
  if (pos + 7U > in.size() || in[pos] != '_' || in[pos + 1U] != 'x' || in[pos + 6U] != '_') {
    return false;
  }
  std::uint16_t value = 0;
  for (std::size_t i = pos + 2U; i < pos + 6U; ++i) {
    const char c = in[i];
    std::uint16_t digit = 0;
    if (c >= '0' && c <= '9') {
      digit = static_cast<std::uint16_t>(c - '0');
    } else if (c >= 'A' && c <= 'F') {
      digit = static_cast<std::uint16_t>(c - 'A' + 10);
    } else if (c >= 'a' && c <= 'f') {
      digit = static_cast<std::uint16_t>(c - 'a' + 10);
    } else {
      return false;
    }
    value = static_cast<std::uint16_t>((value << 4U) | digit);
  }
  *code = value;
  return true;
}

void AppendOoxmlEscape(std::string& out, unsigned char byte) {
  constexpr char kHex[] = "0123456789ABCDEF";
  out.append("_x00");
  out.push_back(kHex[(byte >> 4U) & 0x0FU]);
  out.push_back(kHex[byte & 0x0FU]);
  out.push_back('_');
}

void AppendUtf8(std::string& out, std::uint32_t code_point) {
  if (code_point <= 0x7FU) {
    out.push_back(static_cast<char>(code_point));
  } else if (code_point <= 0x7FFU) {
    out.push_back(static_cast<char>(0xC0U | (code_point >> 6U)));
    out.push_back(static_cast<char>(0x80U | (code_point & 0x3FU)));
  } else if (code_point <= 0xFFFFU) {
    out.push_back(static_cast<char>(0xE0U | (code_point >> 12U)));
    out.push_back(static_cast<char>(0x80U | ((code_point >> 6U) & 0x3FU)));
    out.push_back(static_cast<char>(0x80U | (code_point & 0x3FU)));
  } else {
    out.push_back(static_cast<char>(0xF0U | (code_point >> 18U)));
    out.push_back(static_cast<char>(0x80U | ((code_point >> 12U) & 0x3FU)));
    out.push_back(static_cast<char>(0x80U | ((code_point >> 6U) & 0x3FU)));
    out.push_back(static_cast<char>(0x80U | (code_point & 0x3FU)));
  }
}

std::size_t ValidUtf8SequenceLength(std::string_view in, std::size_t pos) {
  const auto byte_at = [&in](std::size_t i) { return static_cast<unsigned char>(in[i]); };
  const unsigned char first = byte_at(pos);
  const auto continuation = [&byte_at, &in](std::size_t i) {
    return i < in.size() && byte_at(i) >= 0x80U && byte_at(i) <= 0xBFU;
  };
  if (first >= 0xC2U && first <= 0xDFU) {
    return continuation(pos + 1U) ? 2U : 0U;
  }
  if (first == 0xE0U) {
    return pos + 2U < in.size() && byte_at(pos + 1U) >= 0xA0U && byte_at(pos + 1U) <= 0xBFU && continuation(pos + 2U)
               ? 3U
               : 0U;
  }
  if (first >= 0xE1U && first <= 0xECU) {
    return continuation(pos + 1U) && continuation(pos + 2U) ? 3U : 0U;
  }
  if (first == 0xEDU) {
    return pos + 2U < in.size() && byte_at(pos + 1U) >= 0x80U && byte_at(pos + 1U) <= 0x9FU && continuation(pos + 2U)
               ? 3U
               : 0U;
  }
  if (first >= 0xEEU && first <= 0xEFU) {
    return continuation(pos + 1U) && continuation(pos + 2U) ? 3U : 0U;
  }
  if (first == 0xF0U) {
    return pos + 3U < in.size() && byte_at(pos + 1U) >= 0x90U && byte_at(pos + 1U) <= 0xBFU && continuation(pos + 2U) &&
                   continuation(pos + 3U)
               ? 4U
               : 0U;
  }
  if (first >= 0xF1U && first <= 0xF3U) {
    return continuation(pos + 1U) && continuation(pos + 2U) && continuation(pos + 3U) ? 4U : 0U;
  }
  if (first == 0xF4U) {
    return pos + 3U < in.size() && byte_at(pos + 1U) >= 0x80U && byte_at(pos + 1U) <= 0x8FU && continuation(pos + 2U) &&
                   continuation(pos + 3U)
               ? 4U
               : 0U;
  }
  return 0U;
}

void AppendXmlEscapedImpl(std::string& out, std::string_view in, bool attribute) {
  for (std::size_t i = 0; i < in.size(); ++i) {
    const char raw = in[i];
    const unsigned char byte = static_cast<unsigned char>(raw);
    if (byte >= 0x80U) {
      const std::size_t sequence_length = ValidUtf8SequenceLength(in, i);
      if (sequence_length == 0U) {
        // The XML declaration promises UTF-8. Replace an invalid source byte
        // rather than serialising a malformed OOXML part that no conforming
        // XML reader can reopen.
        out.append("\xEF\xBF\xBD");
      } else {
        out.append(in.data() + i, sequence_length);
        i += sequence_length - 1U;
      }
      continue;
    }
    std::uint16_t ignored = 0;
    if (IsOoxmlEscape(in, i, &ignored)) {
      out.append("_x005F_");
      continue;
    }
    if (byte < 0x20U) {
      AppendOoxmlEscape(out, byte);
      continue;
    }
    switch (raw) {
      case '&':
        out.append("&amp;");
        break;
      case '<':
        out.append("&lt;");
        break;
      case '>':
        out.append("&gt;");
        break;
      case '"':
        out.append("&quot;");
        break;
      case '\'':
        out.append("&apos;");
        break;
      case '\t':
        if (attribute) {
          out.append("&#9;");
        } else {
          out.push_back(raw);
        }
        break;
      case '\n':
        if (attribute) {
          out.append("&#10;");
        } else {
          out.push_back(raw);
        }
        break;
      case '\r':
        if (attribute) {
          out.append("&#13;");
        } else {
          out.push_back(raw);
        }
        break;
      default:
        out.push_back(raw);
        break;
    }
  }
}

}  // namespace

void AppendXmlEscaped(std::string& out, std::string_view in) {
  AppendXmlEscapedImpl(out, in, false);
}

void AppendXmlAttrEscaped(std::string& out, std::string_view in) {
  AppendXmlEscapedImpl(out, in, true);
}

void AppendOoxmlTextUnescaped(std::string& out, std::string_view in) {
  for (std::size_t i = 0; i < in.size();) {
    std::uint16_t code = 0;
    if (!IsOoxmlEscape(in, i, &code)) {
      out.push_back(in[i++]);
      continue;
    }
    i += 7U;
    if (code >= 0xD800U && code <= 0xDBFFU) {
      std::uint16_t low = 0;
      if (IsOoxmlEscape(in, i, &low) && low >= 0xDC00U && low <= 0xDFFFU) {
        i += 7U;
        AppendUtf8(out, 0x10000U + ((static_cast<std::uint32_t>(code) - 0xD800U) << 10U) +
                            (static_cast<std::uint32_t>(low) - 0xDC00U));
        continue;
      }
    }
    AppendUtf8(out, code);
  }
}

}  // namespace io
}  // namespace formulon
