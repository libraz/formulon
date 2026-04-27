// Copyright 2026 libraz. Licensed under the MIT License.

#include "io/xml_escape.h"

#include <string>
#include <string_view>

namespace formulon {
namespace io {

void AppendXmlEscaped(std::string& out, std::string_view in) {
  for (char raw : in) {
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
      default:
        out.push_back(raw);
        break;
    }
  }
}

}  // namespace io
}  // namespace formulon
