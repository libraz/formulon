// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

#include "io/xml_escape.h"

#include <string>
#include <string_view>

namespace formulon {
namespace io {

void AppendXmlEscaped(std::string& out, std::string_view in) {
  for (char raw : in) {
    // XML 1.0 forbids U+0000 and U+0001..U+001F (except TAB, LF, CR) in
    // document content. Emitting them verbatim makes the resulting part
    // un-reparseable by pugixml on round-trip. Strip them defensively;
    // Excel itself essentially never produces these in cell text.
    const unsigned char byte = static_cast<unsigned char>(raw);
    if (byte < 0x20U && raw != '\t' && raw != '\n' && raw != '\r') {
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
      default:
        out.push_back(raw);
        break;
    }
  }
}

}  // namespace io
}  // namespace formulon
