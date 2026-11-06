#include "parser/external_book_spelling.h"

#include <cstddef>
#include <cstdint>
#include <string>

#include "parser/ast_format.h"
#include "parser/formula_prefix.h"
#include "parser/parser_detail.h"
#include "parser/reference.h"
#include "utils/strings.h"

namespace formulon {
namespace parser {
namespace {

constexpr std::size_t kMaxIndexDigits = 9;

int HexValue(char c) noexcept {
  if (strings::is_ascii_digit(c)) {
    return c - '0';
  }
  if (c >= 'A' && c <= 'F') {
    return c - 'A' + 10;
  }
  if (c >= 'a' && c <= 'f') {
    return c - 'a' + 10;
  }
  return -1;
}

bool IsSeparator(char c) noexcept {
  return c == '/' || c == '\\';
}

// `s` up to and including its last separator; empty when it has none.
std::string_view DirectoryPrefix(std::string_view s) noexcept {
  const std::size_t pos = s.find_last_of("/\\");
  return pos == std::string_view::npos ? std::string_view() : s.substr(0, pos + 1);
}

std::string WithBackslashes(std::string_view s) {
  std::string out(s);
  for (char& c : out) {
    if (c == '/') {
      c = '\\';
    }
  }
  return out;
}

struct Qualifier {
  std::uint32_t index = 0;
  std::string_view sheet;      // empty for a book-scope name
  std::string_view sheet_end;  // empty unless 3-D
};

// Appends the formula-bar spelling of a qualifier, including its `!`.
void AppendSpelled(std::string* out, const Qualifier& q, const ExternalBookResolver& resolver) {
  ExternalBookDisplay link;
  if (!resolver.resolve(resolver.ctx, q.index, &link)) {
    link.path.clear();
    link.book = std::to_string(q.index);
  }
  const bool book_scope = q.sheet.empty();
  bool quote = !link.path.empty() || external_book_needs_quoting(link.book) ||
               (!book_scope && (external_sheet_needs_quoting(q.sheet) || !q.sheet_end.empty()));
  // `Src2!Name` reads as a local sheet's name, so it takes that quoting.
  if (book_scope && !quote) {
    quote = local_sheet_needs_quoting_a1(link.book);
  }
  if (quote) {
    out->push_back('\'');
    append_quote_escaped(link.path, *out);
  }
  if (book_scope) {
    append_quote_escaped(link.book, *out);
  } else {
    out->push_back('[');
    if (quote) {
      append_quote_escaped(link.book, *out);
    } else {
      out->append(link.book);
    }
    out->push_back(']');
    if (quote) {
      append_quote_escaped(q.sheet, *out);
    } else {
      out->append(q.sheet);
    }
    if (!q.sheet_end.empty()) {
      out->push_back(':');
      append_quote_escaped(q.sheet_end, *out);
    }
  }
  if (quote) {
    out->push_back('\'');
  }
  out->push_back('!');
}

// Reads `[digits]` at `pos`; returns the index after `]`, or npos.
std::size_t ParseBracketIndex(std::string_view s, std::size_t pos, std::uint32_t* index) noexcept {
  if (pos >= s.size() || s[pos] != '[') {
    return std::string_view::npos;
  }
  std::size_t i = pos + 1;
  std::uint32_t value = 0;
  while (i < s.size() && strings::is_ascii_digit(s[i])) {
    value = value * 10 + static_cast<std::uint32_t>(s[i] - '0');
    ++i;
  }
  if (i == pos + 1 || i - pos - 1 > kMaxIndexDigits || i >= s.size() || s[i] != ']') {
    return std::string_view::npos;
  }
  *index = value;
  return i + 1;
}

std::size_t ScanRun(std::string_view s, std::size_t pos) noexcept {
  while (pos < s.size() && formula_prefix_is_ident_byte(s[pos])) {
    ++pos;
  }
  return pos;
}

// Parses an unquoted qualifier at `pos` (`[N]Sheet!`, `[N]S1:S2!`, `[N]!`).
// Returns the index after its `!`, or npos.
std::size_t ParseBareQualifier(std::string_view s, std::size_t pos, Qualifier* q) noexcept {
  std::size_t i = ParseBracketIndex(s, pos, &q->index);
  if (i == std::string_view::npos) {
    return i;
  }
  if (i < s.size() && s[i] == '!') {
    q->sheet = std::string_view();
    q->sheet_end = std::string_view();
    return q->index == 0 ? std::string_view::npos : i + 1;
  }
  std::size_t end = ScanRun(s, i);
  if (end == i) {
    return std::string_view::npos;
  }
  q->sheet = s.substr(i, end - i);
  q->sheet_end = std::string_view();
  if (end < s.size() && s[end] == ':') {
    const std::size_t second = ScanRun(s, end + 1);
    if (second == end + 1) {
      return std::string_view::npos;
    }
    q->sheet_end = s.substr(end + 1, second - end - 1);
    end = second;
  }
  if (end >= s.size() || s[end] != '!') {
    return std::string_view::npos;
  }
  return q->index == 0 ? std::string_view::npos : end + 1;
}

// Parses a quoted qualifier whose opening quote is at `pos`. The content is
// un-doubled into `content`, which `q`'s views point into. Returns the index
// after the `!`, or npos when it is not an indexed external qualifier.
std::size_t ParseQuotedQualifier(std::string_view s, std::size_t pos, std::string* content, Qualifier* q) {
  content->clear();
  std::size_t i = pos + 1;
  for (;; ++i) {
    if (i >= s.size()) {
      return std::string_view::npos;
    }
    if (s[i] == '\'') {
      if (i + 1 < s.size() && s[i + 1] == '\'') {
        ++i;
      } else {
        break;
      }
    }
    content->push_back(s[i]);
  }
  const std::size_t after = i + 1;
  if (after >= s.size() || s[after] != '!') {
    return std::string_view::npos;
  }
  const std::string_view body(*content);
  const std::size_t rest = ParseBracketIndex(body, 0, &q->index);
  if (rest == std::string_view::npos || rest >= body.size() || q->index == 0) {
    return std::string_view::npos;
  }
  std::string_view sheets = body.substr(rest);
  if (sheets.find('[') != std::string_view::npos) {
    return std::string_view::npos;
  }
  const std::size_t colon = sheets.find(':');
  if (colon == std::string_view::npos) {
    q->sheet = sheets;
    q->sheet_end = std::string_view();
  } else {
    q->sheet = sheets.substr(0, colon);
    q->sheet_end = sheets.substr(colon + 1);
    if (q->sheet.empty() || q->sheet_end.empty()) {
      return std::string_view::npos;
    }
  }
  return after + 1;
}

// The index just past the `quote`-delimited run opening at `pos`, in which a
// doubled quote is literal; the end of `s` when the run is unterminated.
std::size_t QuotedRunEnd(std::string_view s, std::size_t pos, char quote) noexcept {
  std::size_t j = pos + 1;
  while (j < s.size()) {
    if (s[j] == quote) {
      if (j + 1 < s.size() && s[j + 1] == quote) {
        ++j;
      } else {
        break;
      }
    }
    ++j;
  }
  return j < s.size() ? j + 1 : s.size();
}

// True when a defined name (not a cell, area, number or boolean) starts at
// `pos`.
bool NameStartsAt(std::string_view s, std::size_t pos) noexcept {
  const std::size_t end = ScanRun(s, pos);
  return end > pos && !local_sheet_needs_quoting_a1(s.substr(pos, end - pos)) && (end >= s.size() || s[end] != ':');
}

// True when `c` before a `[` or a bare run makes it a continuation rather
// than the start of an operand.
bool ContinuesOperand(char c) noexcept {
  return formula_prefix_is_ident_byte(c) || c == ']' || c == ')' || c == '\'';
}

// Appends the local spelling of the sheet part `sheets` (`S` or `S1:S2`)
// followed by `!`.
void AppendLocalQualifier(std::string* out, std::string_view sheets) {
  bool quote = false;
  std::size_t begin = 0;
  while (begin <= sheets.size()) {
    std::size_t colon = sheets.find(':', begin);
    if (colon == std::string_view::npos) {
      colon = sheets.size();
    }
    quote = quote || local_sheet_needs_quoting_a1(sheets.substr(begin, colon - begin));
    begin = colon + 1;
  }
  append_sheet_name_quoted_if(sheets, quote, *out);
  out->push_back('!');
}

// `[book]` opening `content` when it names `file_name`: returns the index
// after `]`, else npos.
std::size_t SelfBookBracketEnd(std::string_view content, std::size_t pos, std::string_view file_name) noexcept {
  if (pos >= content.size() || content[pos] != '[') {
    return std::string_view::npos;
  }
  const std::size_t close = content.find(']', pos);
  if (close == std::string_view::npos ||
      !strings::case_insensitive_eq(content.substr(pos + 1, close - pos - 1), file_name)) {
    return std::string_view::npos;
  }
  return close + 1;
}

}  // namespace

std::string percent_decode(std::string_view s) {
  std::string out;
  out.reserve(s.size());
  for (std::size_t i = 0; i < s.size(); ++i) {
    if (s[i] == '%' && i + 2 < s.size() && HexValue(s[i + 1]) >= 0 && HexValue(s[i + 2]) >= 0) {
      out.push_back(static_cast<char>(HexValue(s[i + 1]) * 16 + HexValue(s[i + 2])));
      i += 2;
    } else {
      out.push_back(s[i]);
    }
  }
  return out;
}

std::string spell_external_books(std::string_view stored, const ExternalBookResolver& resolver) {
  if (stored.find('[') == std::string_view::npos) {
    return std::string(stored);
  }
  std::string out;
  out.reserve(stored.size() + 32);
  std::string content;
  int depth = 0;
  std::size_t i = 0;
  while (i < stored.size()) {
    const char c = stored[i];
    if (c == '"') {
      const std::size_t j = QuotedRunEnd(stored, i, '"');
      out.append(stored.substr(i, j - i));
      i = j;
      continue;
    }
    if (c == '\'' && depth == 0) {
      Qualifier q;
      const std::size_t next = ParseQuotedQualifier(stored, i, &content, &q);
      if (next != std::string_view::npos) {
        AppendSpelled(&out, q, resolver);
        i = next;
        continue;
      }
      // Copy the quoted name whole so brackets inside it are not scanned.
      const std::size_t j = QuotedRunEnd(stored, i, '\'');
      out.append(stored.substr(i, j - i));
      i = j;
      continue;
    }
    if (c == '[') {
      const bool operand_position = depth == 0 && (i == 0 || !ContinuesOperand(stored[i - 1]));
      if (operand_position) {
        Qualifier q;
        const std::size_t next = ParseBareQualifier(stored, i, &q);
        if (next != std::string_view::npos) {
          AppendSpelled(&out, q, resolver);
          i = next;
          continue;
        }
      }
      ++depth;
    } else if (c == ']' && depth > 0) {
      --depth;
    }
    out.push_back(c);
    ++i;
  }
  return out;
}

std::string spell_self_book_names(std::string_view text, std::string_view file_name) {
  if (file_name.empty() || text.find("[0]!") == std::string_view::npos) {
    return std::string(text);
  }
  std::string book;
  append_sheet_name_quoted_if(file_name, self_book_name_needs_quoting(file_name), book);
  std::string out;
  out.reserve(text.size() + book.size());
  std::size_t i = 0;
  while (i < text.size()) {
    const char c = text[i];
    if (c == '"' || c == '\'') {
      const std::size_t j = QuotedRunEnd(text, i, c);
      out.append(text.substr(i, j - i));
      i = j;
      continue;
    }
    if (text.compare(i, 4, "[0]!") == 0 && (i == 0 || !ContinuesOperand(text[i - 1])) && NameStartsAt(text, i + 4)) {
      out.append(book);
      out.push_back('!');
      i += 4;
      continue;
    }
    out.push_back(c);
    ++i;
  }
  return out;
}

std::string bind_self_book_names(std::string_view text, std::string_view file_name) {
  if (file_name.empty()) {
    return std::string(text);
  }
  std::string out;
  out.reserve(text.size());
  std::string content;
  int depth = 0;
  std::size_t i = 0;
  while (i < text.size()) {
    const char c = text[i];
    if (c == '"') {
      const std::size_t j = QuotedRunEnd(text, i, '"');
      out.append(text.substr(i, j - i));
      i = j;
      continue;
    }
    if (c == '\'' && depth == 0) {
      const std::size_t j = QuotedRunEnd(text, i, '\'');
      const bool closed = j - i >= 2 && text[j - 1] == '\'';
      content.clear();
      for (std::size_t k = i + 1; closed && k + 1 < j; ++k) {
        content.push_back(text[k]);
        if (text[k] == '\'') {
          ++k;
        }
      }
      if (closed && j < text.size() && text[j] == '!') {
        if (strings::case_insensitive_eq(content, file_name) && NameStartsAt(text, j + 1)) {
          out.append("[0]!");
          i = j + 1;
          continue;
        }
        const std::size_t sheets = SelfBookBracketEnd(content, 0, file_name);
        if (sheets != std::string_view::npos && sheets < content.size()) {
          AppendLocalQualifier(&out, std::string_view(content).substr(sheets));
          i = j + 1;
          continue;
        }
      }
      out.append(text.substr(i, j - i));
      i = j;
      continue;
    }
    if (depth == 0 && c == '[' && (i == 0 || !ContinuesOperand(text[i - 1]))) {
      const std::size_t sheets_begin = SelfBookBracketEnd(text, i, file_name);
      if (sheets_begin != std::string_view::npos) {
        std::size_t end = ScanRun(text, sheets_begin);
        if (end > sheets_begin && end < text.size() && text[end] == ':') {
          const std::size_t second = ScanRun(text, end + 1);
          end = second > end + 1 ? second : end;
        }
        if (end > sheets_begin && end < text.size() && text[end] == '!') {
          AppendLocalQualifier(&out, text.substr(sheets_begin, end - sheets_begin));
          i = end + 1;
          continue;
        }
      }
    }
    if (depth == 0 && formula_prefix_is_ident_byte(c)) {
      const std::size_t end = ScanRun(text, i);
      if (end < text.size() && text[end] == '!' && strings::case_insensitive_eq(text.substr(i, end - i), file_name) &&
          NameStartsAt(text, end + 1)) {
        out.append("[0]!");
        i = end + 1;
        continue;
      }
      out.append(text.substr(i, end - i));
      i = end;
      continue;
    }
    if (c == '[') {
      ++depth;
    } else if (c == ']' && depth > 0) {
      --depth;
    }
    out.push_back(c);
    ++i;
  }
  return out;
}

std::string display_path_for_link_target(std::string_view target) {
  if (strings::starts_with(target, "file://")) {
    const std::string_view rest = target.substr(7);
    if (!rest.empty() && rest.front() == '/') {
      // file:///<path>: a drive path or a POSIX absolute path.
      const std::string path = percent_decode(rest.substr(1));
      const bool drive = path.size() >= 3 && strings::is_ascii_alpha(path[0]) && path[1] == ':' && IsSeparator(path[2]);
      if (drive) {
        return WithBackslashes(DirectoryPrefix(path));
      }
      return std::string(DirectoryPrefix(percent_decode(rest)));
    }
    // file://server/share/...: a UNC path.
    const std::string path = percent_decode(rest);
    const std::string_view dir = DirectoryPrefix(path);
    if (path.empty()) {
      return std::string();
    }
    return "\\\\" + WithBackslashes(dir.empty() ? std::string_view(path) : dir) + (dir.empty() ? "\\" : "");
  }
  if (strings::starts_with(target, "http://") || strings::starts_with(target, "https://")) {
    const std::size_t host = target.find("//") + 2;
    const std::size_t last = target.find_last_of('/');
    return last >= host ? std::string(target.substr(0, last + 1)) : std::string();
  }
  if (!target.empty() && target.front() == '/') {
    return std::string(DirectoryPrefix(target));
  }
  return std::string();
}

}  // namespace parser
}  // namespace formulon
