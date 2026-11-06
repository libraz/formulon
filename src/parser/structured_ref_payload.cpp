//
// Structured-reference payload grammar and spellings. See
// `structured_ref_payload.h` for the contract.

#include "parser/structured_ref_payload.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "parser/parser_detail.h"
#include "utils/strings.h"

namespace formulon {
namespace parser {
namespace {

using Specs = StructuredRefSpecifiers;

// Storage order of the specifier items (`[#Headers],[#Data]`).
struct SpecWord {
  std::uint8_t bit;
  std::string_view word;
};
constexpr SpecWord kSpecWords[] = {
    {Specs::kAll, "#All"},       {Specs::kHeaders, "#Headers"},  {Specs::kData, "#Data"},
    {Specs::kTotals, "#Totals"}, {Specs::kThisRow, "#This Row"},
};

std::uint8_t ClassifySpecifier(std::string_view text) noexcept {
  for (const SpecWord& spec : kSpecWords) {
    if (strings::case_insensitive_eq(text, spec.word)) {
      return spec.bit;
    }
  }
  return 0U;
}

void SkipSpaces(std::string_view s, std::size_t* i) noexcept {
  while (*i < s.size() && (s[*i] == ' ' || s[*i] == '\t')) {
    ++*i;
  }
}

// Unescapes `s` up to an unescaped `]` (when `in_item`) or the end. An
// unescaped `[`, or `]` outside an item, is malformed.
bool Unescape(std::string_view s, std::size_t* i, bool in_item, std::string* out) {
  for (; *i < s.size(); ++*i) {
    const char c = s[*i];
    if (c == '\'') {
      if (++*i == s.size()) {
        return false;
      }
      out->push_back(s[*i]);
    } else if (c == ']') {
      return in_item;
    } else if (c == '[') {
      return false;
    } else {
      out->push_back(c);
    }
  }
  return !in_item;
}

// Reads `[item]` at `*i`; `*spec` is its specifier bit, 0 for a column.
bool ReadItem(std::string_view s, std::size_t* i, std::string* column, std::uint8_t* spec) {
  SkipSpaces(s, i);
  if (*i >= s.size() || s[*i] != '[') {
    return false;
  }
  const std::size_t start = ++*i;
  if (!Unescape(s, i, /*in_item=*/true, column)) {
    return false;
  }
  const std::string_view raw = s.substr(start, *i - start);
  ++*i;  // `]`
  *spec = 0U;
  if (!raw.empty() && raw.front() == '#') {
    *spec = ClassifySpecifier(strings::trim_blanks(raw));
    return *spec != 0U;
  }
  return !column->empty();
}

// Reads `[A]` or `[A]:[B]` into `out->columns`.
bool ReadColumnGroup(std::string_view s, std::size_t* i, std::string first, StructuredRefPayload* out) {
  out->columns.push_back(std::move(first));
  SkipSpaces(s, i);
  if (*i < s.size() && s[*i] == ':') {
    ++*i;
    std::string last;
    std::uint8_t spec = 0U;
    if (!ReadItem(s, i, &last, &spec) || spec != 0U) {
      return false;
    }
    out->columns.push_back(std::move(last));
  }
  return true;
}

bool ParseItemList(std::string_view s, StructuredRefPayload* out) {
  std::size_t i = 0;
  while (true) {
    std::string column;
    std::uint8_t spec = 0U;
    if (!ReadItem(s, &i, &column, &spec)) {
      return false;
    }
    if (spec != 0U) {
      out->specifiers |= spec;
    } else if (!out->columns.empty() || !ReadColumnGroup(s, &i, std::move(column), out)) {
      return false;
    }
    SkipSpaces(s, &i);
    if (i == s.size()) {
      return true;
    }
    if (s[i] != ',') {
      return false;
    }
    ++i;
  }
}

bool HasEdgeSpace(std::string_view name) noexcept {
  return !name.empty() && (name.front() == ' ' || name.back() == ' ');
}

// `name` with `'` before each of `' [ ] #`, and before `@` for display.
void AppendEscaped(std::string_view name, bool display, std::string& out) {
  for (const char c : name) {
    if (c == '\'' || c == '[' || c == ']' || c == '#' || (display && c == '@')) {
      out.push_back('\'');
    }
    out.push_back(c);
  }
}

void AppendBracketed(std::string_view name, bool display, std::string& out) {
  out.push_back('[');
  AppendEscaped(name, display, out);
  out.push_back(']');
}

// `[A]` or `[A]:[B]`.
void AppendColumns(const std::vector<std::string>& columns, bool display, std::string& out) {
  for (std::size_t c = 0; c < columns.size(); ++c) {
    if (c > 0U) {
      out.push_back(':');
    }
    AppendBracketed(columns[c], display, out);
  }
}

// The long-form item list: specifiers in storage order, then the columns.
void AppendItemList(const StructuredRefPayload& p, bool display, std::string& out) {
  bool first = true;
  for (const SpecWord& spec : kSpecWords) {
    if ((p.specifiers & spec.bit) != 0U) {
      out.append(first ? "[" : ",[").append(spec.word).push_back(']');
      first = false;
    }
  }
  if (!first && !p.columns.empty()) {
    out.push_back(',');
  }
  AppendColumns(p.columns, display, out);
}

// After `@`, a name of ASCII letters, digits, `|` and non-ASCII characters
// needs no brackets.
bool BareAfterAt(std::string_view name) noexcept {
  for (const char c : name) {
    if (!strings::is_ascii_alpha(c) && !strings::is_ascii_digit(c) && c != '|' &&
        static_cast<unsigned char>(c) < 0x80U) {
      return false;
    }
  }
  return !name.empty();
}

bool SingleBit(std::uint8_t bits) noexcept {
  return bits != 0U && (bits & (bits - 1U)) == 0U;
}

}  // namespace

std::optional<StructuredRefPayload> parse_structured_ref_payload_text(std::string_view payload) {
  const std::string_view p = strings::trim_blanks(payload);
  StructuredRefPayload out;
  if (p.empty()) {
    return out;
  }
  std::size_t i = 0;
  if (p.front() == '@') {
    out.specifiers = Specs::kThisRow;
    i = 1;
    SkipSpaces(p, &i);
    if (i == p.size()) {
      return out;
    }
    if (p[i] == '[') {
      std::string first;
      std::uint8_t spec = 0U;
      if (!ReadItem(p, &i, &first, &spec) || spec != 0U || !ReadColumnGroup(p, &i, std::move(first), &out)) {
        return std::nullopt;
      }
      SkipSpaces(p, &i);
      return i == p.size() ? std::optional<StructuredRefPayload>(std::move(out)) : std::nullopt;
    }
  } else if (p.front() == '#') {
    out.specifiers = ClassifySpecifier(p);
    return out.specifiers != 0U ? std::optional<StructuredRefPayload>(std::move(out)) : std::nullopt;
  } else if (p.front() == '[') {
    return ParseItemList(p, &out) ? std::optional<StructuredRefPayload>(std::move(out)) : std::nullopt;
  }
  std::string column;
  if (!Unescape(p, &i, /*in_item=*/false, &column)) {
    return std::nullopt;
  }
  out.columns.push_back(std::move(column));
  return out;
}

std::string stored_payload(const StructuredRefPayload& payload) {
  std::string out;
  if (payload.specifiers == 0U && payload.columns.size() == 1U) {
    const std::string& name = payload.columns.front();
    // A leading `@` would read back as `#This Row`.
    if (HasEdgeSpace(name) || (!name.empty() && name.front() == '@')) {
      AppendBracketed(name, /*display=*/false, out);
    } else {
      AppendEscaped(name, /*display=*/false, out);
    }
    return out;
  }
  if (payload.columns.empty() && SingleBit(payload.specifiers)) {
    for (const SpecWord& spec : kSpecWords) {
      if (payload.specifiers == spec.bit) {
        out.append(spec.word);
      }
    }
    return out;
  }
  AppendItemList(payload, /*display=*/false, out);
  return out;
}

std::string display_structured_ref(std::string_view table, const StructuredRefPayload& payload, bool inside_own_table) {
  const std::size_t cols = payload.columns.size();
  const std::uint8_t specs = payload.specifiers;
  std::string out;
  if (specs == 0U && cols == 0U) {
    out.append(table);
    return out;
  }
  if (!(inside_own_table && cols == 1U && (specs == 0U || specs == Specs::kThisRow))) {
    out.append(table);
  }
  out.push_back('[');
  if (specs == Specs::kThisRow) {
    out.push_back('@');
    if (cols == 1U && BareAfterAt(payload.columns.front())) {
      out.append(payload.columns.front());
    } else if (cols > 0U) {
      AppendColumns(payload.columns, /*display=*/true, out);
    }
  } else if (specs == 0U && cols == 1U && !HasEdgeSpace(payload.columns.front())) {
    AppendEscaped(payload.columns.front(), /*display=*/true, out);
  } else if (cols == 0U && SingleBit(specs)) {
    out.append(stored_payload(payload));
  } else {
    AppendItemList(payload, /*display=*/true, out);
  }
  out.push_back(']');
  return out;
}

}  // namespace parser
}  // namespace formulon
