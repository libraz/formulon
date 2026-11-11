#include "hyperlink_address.h"

#include <cstddef>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

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

bool is_ascii_alpha(char c) {
  return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
}

bool is_separator(char c) {
  return c == '/' || c == '\\';
}

/// Whether `text` opens with a drive letter and a separator (`C:\`, `C:/`).
bool is_drive_path(std::string_view text) {
  return text.size() >= 3U && is_ascii_alpha(text[0]) && text[1] == ':' && is_separator(text[2]);
}

/// Whether `text` opens with a URL scheme of two or more characters; a
/// single letter before `:` is a drive.
bool has_scheme(std::string_view text) {
  if (text.empty() || !is_ascii_alpha(text[0])) {
    return false;
  }
  for (std::size_t i = 1; i < text.size(); ++i) {
    const char c = text[i];
    if (c == ':') {
      return i >= 2U;
    }
    if (!is_ascii_alpha(c) && !(c >= '0' && c <= '9') && c != '+' && c != '-' && c != '.') {
      return false;
    }
  }
  return false;
}

std::string decode_spaces(std::string_view text) {
  std::string out;
  out.reserve(text.size());
  for (std::size_t i = 0; i < text.size(); ++i) {
    if (text[i] == '%' && i + 2 < text.size() && text[i + 1] == '2' && text[i + 2] == '0') {
      out.push_back(' ');
      i += 2;
      continue;
    }
    out.push_back(text[i]);
  }
  return out;
}

bool same_ignore_case(std::string_view a, std::string_view b) {
  if (a.size() != b.size()) {
    return false;
  }
  for (std::size_t i = 0; i < a.size(); ++i) {
    if (ascii_lower(a[i]) != ascii_lower(b[i])) {
      return false;
    }
  }
  return true;
}

enum class RootKind { kPosix, kDrive, kUnc };

/// An absolute local path: the root (empty for POSIX, the drive letter, or
/// `server/share`) and the segments below it.
struct LocalPath {
  RootKind kind = RootKind::kPosix;
  std::string root;
  std::vector<std::string> segments;
};

/// Appends the non-empty pieces of `text` split at `/`, and at `\` too when
/// `backslash` is set.
void split_segments(std::string_view text, bool backslash, std::vector<std::string>& out) {
  std::size_t start = 0;
  for (std::size_t i = 0; i <= text.size(); ++i) {
    if (i == text.size() || text[i] == '/' || (backslash && text[i] == '\\')) {
      if (i > start) {
        out.emplace_back(text.substr(start, i - start));
      }
      start = i + 1U;
    }
  }
}

/// `path` without a scheme: `/a/b`, `C:\a` / `C:/a`, `\\server\share\a` or
/// `//server/share/a`. `\` separates only in the drive and UNC forms.
std::optional<LocalPath> parse_plain_path(std::string_view path) {
  LocalPath out;
  if (path.size() >= 2U && ((path[0] == '/' && path[1] == '/') || (path[0] == '\\' && path[1] == '\\'))) {
    std::vector<std::string> pieces;
    split_segments(path.substr(2), /*backslash=*/true, pieces);
    if (pieces.size() < 2U) {
      return std::nullopt;
    }
    out.kind = RootKind::kUnc;
    out.root = pieces[0] + "/" + pieces[1];
    out.segments.assign(pieces.begin() + 2, pieces.end());
    return out;
  }
  if (is_drive_path(path)) {
    out.kind = RootKind::kDrive;
    out.root = std::string(1, path[0]);
    split_segments(path.substr(3), /*backslash=*/true, out.segments);
    return out;
  }
  if (!path.empty() && path[0] == '/') {
    split_segments(path.substr(1), /*backslash=*/false, out.segments);
    return out;
  }
  return std::nullopt;
}

/// A local absolute target: a `file:` URL with an empty host (`file:///a`,
/// `file:///C:/a`, `file:////server/share/a`), whose `%20` reads as a space,
/// or a plain absolute path, whose `%` is literal.
std::optional<LocalPath> parse_local_absolute(std::string_view target) {
  if (!starts_with_ignore_case(target, "file:")) {
    return has_scheme(target) ? std::nullopt : parse_plain_path(target);
  }
  const std::string path = decode_spaces(target.substr(5));
  if (path.compare(0, 4, "////") == 0) {
    return parse_plain_path(std::string_view(path).substr(2));
  }
  if (path.compare(0, 3, "///") != 0) {
    return std::nullopt;
  }
  const std::string_view rest = std::string_view(path).substr(2);
  if (is_drive_path(rest.substr(1))) {
    return parse_plain_path(rest.substr(1));
  }
  return parse_plain_path(rest);
}

std::string file_url(const LocalPath& path) {
  std::string out;
  switch (path.kind) {
    case RootKind::kPosix:
      out = "file:///";
      break;
    case RootKind::kDrive:
      out = "file:///" + path.root + ":/";
      break;
    case RootKind::kUnc:
      out = "file:////" + path.root + (path.segments.empty() ? "" : "/");
      break;
  }
  for (std::size_t i = 0; i < path.segments.size(); ++i) {
    if (i != 0U) {
      out.push_back('/');
    }
    out += path.segments[i];
  }
  return out;
}

/// Whether `target` is a path relative to the book: no scheme, root, drive
/// or leading `#`.
bool is_relative_local(std::string_view target) {
  if (target.empty() || is_separator(target[0]) || target[0] == '#') {
    return false;
  }
  if (target.size() >= 2U && is_ascii_alpha(target[0]) && target[1] == ':') {
    return false;
  }
  return !has_scheme(target);
}

}  // namespace

NormalizedHyperlinkTarget normalize_hyperlink_target(std::string_view input) {
  NormalizedHyperlinkTarget out;
  const bool unc = input.size() >= 2U && input[0] == '\\' && input[1] == '\\';
  if (is_drive_path(input) || unc || (!input.empty() && input.front() == '/')) {
    std::string_view path = input;
    if (const std::size_t hash = path.find('#'); hash != std::string_view::npos) {
      out.fragment = std::string(path.substr(hash + 1));
      path = path.substr(0, hash);
    }
    // A POSIX path holding `%` stays a path; only `%20` is decoded.
    const bool posix = path.front() == '/' && !(path.size() >= 2U && path[1] == '/');
    if (posix && path.find('%') != std::string_view::npos) {
      out.target = decode_spaces(path);
      return out;
    }
    std::size_t from = 0;
    if (unc) {
      out.target = "file:////" + std::string(path.substr(2));
      from = 9U;
    } else if (path.front() == '/') {
      out.target = "file://" + std::string(path);
      from = out.target.size();
    } else {
      out.target = "file:///" + std::string(path);
      from = 8U;
    }
    for (std::size_t i = from; i < out.target.size(); ++i) {
      if (out.target[i] == '\\') {
        out.target[i] = '/';
      }
    }
    return out;
  }
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

std::optional<std::string> book_relative_target(std::string_view abs_target, std::string_view directory) {
  if (directory.empty()) {
    return std::nullopt;
  }
  const std::optional<LocalPath> dir = parse_plain_path(directory);
  const std::optional<LocalPath> target = parse_local_absolute(abs_target);
  if (!dir || !target || dir->kind != target->kind || !same_ignore_case(dir->root, target->root) ||
      target->segments.empty()) {
    return std::nullopt;
  }
  const std::size_t folders = target->segments.size() - 1U;
  std::size_t shared = 0;
  while (shared < dir->segments.size() && shared < folders && dir->segments[shared] == target->segments[shared]) {
    ++shared;
  }
  if (shared == 0U) {
    return std::nullopt;
  }
  // A plain path's `%` is literal, so the saved form escapes it.
  const bool plain = !starts_with_ignore_case(abs_target, "file:");
  std::string out;
  for (std::size_t i = shared; i < dir->segments.size(); ++i) {
    out += "../";
  }
  for (std::size_t i = shared; i < target->segments.size(); ++i) {
    if (i != shared) {
      out.push_back('/');
    }
    for (const char c : target->segments[i]) {
      if (plain && c == '%') {
        out += "%25";
      } else {
        out.push_back(c);
      }
    }
  }
  return out;
}

std::optional<std::string> resolve_relative_target(std::string_view target, std::string_view directory) {
  if (directory.empty() || !is_relative_local(target)) {
    return std::nullopt;
  }
  std::optional<LocalPath> path = parse_plain_path(directory);
  if (!path) {
    return std::nullopt;
  }
  std::vector<std::string> pieces;
  split_segments(decode_spaces(target), /*backslash=*/path->kind != RootKind::kPosix, pieces);
  for (std::string& piece : pieces) {
    if (piece == ".") {
      continue;
    }
    if (piece == "..") {
      if (path->segments.empty()) {
        return std::nullopt;
      }
      path->segments.pop_back();
      continue;
    }
    path->segments.push_back(std::move(piece));
  }
  return file_url(*path);
}

SavedHyperlink saved_hyperlink(const Hyperlink& link, std::string_view directory) {
  SavedHyperlink out{link.target, link.display};
  if (std::optional<std::string> relative = book_relative_target(link.target, directory)) {
    out.target = std::move(*relative);
    if (out.display.empty()) {
      out.display = link.location.empty() ? out.target : out.target + " - " + link.location;
    }
  }
  return out;
}

}  // namespace formulon
