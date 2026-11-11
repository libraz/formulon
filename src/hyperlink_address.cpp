#include "hyperlink_address.h"

#include <cstddef>
#include <utility>

namespace formulon {

namespace {

char ascii_lower(char c) {
  return (c >= 'A' && c <= 'Z') ? static_cast<char>(c - 'A' + 'a') : c;
}

bool starts_with_ignore_case(std::string_view text, std::string_view prefix) {
  if (text.size() < prefix.size()) {
    return false;
  }
  for (std::size_t i = 0; i < prefix.size(); ++i) {
    if (ascii_lower(text[i]) != prefix[i]) {
      return false;
    }
  }
  return true;
}

int hex_value(char c) {
  if (c >= '0' && c <= '9') {
    return c - '0';
  }
  if (c >= 'a' && c <= 'f') {
    return c - 'a' + 10;
  }
  if (c >= 'A' && c <= 'F') {
    return c - 'A' + 10;
  }
  return -1;
}

bool is_unreserved(char c) {
  return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '-' || c == '.' ||
         c == '_' || c == '~';
}

/// `text` with every `%XX` that spells an unreserved character decoded.
std::string decode_unreserved(std::string_view text) {
  std::string out;
  out.reserve(text.size());
  for (std::size_t i = 0; i < text.size(); ++i) {
    if (text[i] == '%' && i + 2 < text.size()) {
      const int hi = hex_value(text[i + 1]);
      const int lo = hex_value(text[i + 2]);
      if (hi >= 0 && lo >= 0) {
        const char decoded = static_cast<char>(hi * 16 + lo);
        if (is_unreserved(decoded)) {
          out.push_back(decoded);
          i += 2;
          continue;
        }
      }
    }
    out.push_back(text[i]);
  }
  return out;
}

/// Length of the scheme prefix (`http://`, `https://`, `ftp://`), else 0.
std::size_t hierarchical_scheme_length(std::string_view text) {
  for (const std::string_view scheme :
       {std::string_view("http://"), std::string_view("https://"), std::string_view("ftp://")}) {
    if (starts_with_ignore_case(text, scheme)) {
      return scheme.size();
    }
  }
  return 0U;
}

}  // namespace

NormalizedHyperlinkTarget normalize_hyperlink_target(std::string_view input) {
  NormalizedHyperlinkTarget out;
  if (const std::size_t scheme_len = hierarchical_scheme_length(input); scheme_len != 0U) {
    const std::string_view rest = input.substr(scheme_len);
    const std::size_t authority_end = rest.find_first_of("/?#");
    const std::string_view authority = rest.substr(0, authority_end);
    if (authority.find('@') != std::string_view::npos) {
      out.target = std::string(input);
      return out;
    }
    std::string scheme(input.substr(0, scheme_len));
    for (char& c : scheme) {
      c = ascii_lower(c);
    }
    std::string_view tail = authority_end == std::string_view::npos ? std::string_view{} : rest.substr(authority_end);
    const std::size_t hash = tail.find('#');
    if (hash != std::string_view::npos) {
      out.fragment = std::string(tail.substr(hash + 1));
      tail = tail.substr(0, hash);
    }
    out.target = std::move(scheme);
    for (const char c : authority) {
      out.target.push_back(ascii_lower(c));
    }
    if (tail.empty() || tail.front() == '?') {
      out.target.push_back('/');
    }
    out.target += decode_unreserved(tail);
    return out;
  }
  if (starts_with_ignore_case(input, "mailto:")) {
    out.target = "mailto:" + std::string(input.substr(7));
    return out;
  }
  if (!input.empty() && input.front() == '/') {
    out.target = "file://" + std::string(input);
    return out;
  }
  if (input.size() >= 2U && input[0] == '\\' && input[1] == '\\') {
    out.target = "file:////" + std::string(input.substr(2));
    for (std::size_t i = 9U; i < out.target.size(); ++i) {
      if (out.target[i] == '\\') {
        out.target[i] = '/';
      }
    }
    return out;
  }
  out.target = std::string(input);
  return out;
}

std::string normalize_hyperlink(Hyperlink& link) {
  NormalizedHyperlinkTarget normalized = normalize_hyperlink_target(link.target);
  link.target = std::move(normalized.target);
  std::string derived = link.target;
  if (!normalized.fragment.empty()) {
    derived += " - " + normalized.fragment;
    if (link.location.empty()) {
      link.location = std::move(normalized.fragment);
    }
  }
  return derived;
}

}  // namespace formulon
