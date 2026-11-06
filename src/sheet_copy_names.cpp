#include "sheet_copy_names.h"

#include "sheet_name.h"
#include "utils/a1_ref.h"
#include "utils/strings.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"

namespace formulon {
namespace {

/// Strips a trailing " (<digits>)" and returns the base, or the whole name
/// when nothing would remain or the suffix is absent.
std::string_view strip_copy_suffix(std::string_view name) {
  if (name.size() < 2 || name.back() != ')') {
    return name;
  }
  std::size_t pos = name.size() - 1;
  std::size_t digits_end = pos;
  while (pos > 0 && strings::is_ascii_digit(name[pos - 1])) {
    --pos;
  }
  if (pos == digits_end || pos < 2 || name[pos - 1] != '(' || name[pos - 2] != ' ') {
    return name;
  }
  const std::size_t base_len = pos - 2;
  return base_len == 0 ? name : name.substr(0, base_len);
}

/// Whole-string A1 reference: 1-3 letters up to XFD, 1-7 digits up to 1048576.
bool is_a1_ref(std::string_view s) {
  std::size_t i = 0;
  std::uint32_t col = 0;
  while (i < s.size() && strings::is_ascii_alpha(s[i])) {
    col = col * 26 + static_cast<std::uint32_t>(strings::ascii_to_upper(s[i]) - 'A' + 1);
    ++i;
    if (i > a1::kMaxColumnLetters) {
      return false;
    }
  }
  if (i == 0) {
    return false;
  }
  const std::size_t digits = s.size() - i;
  if (digits == 0 || digits > 7) {
    return false;
  }
  std::uint32_t row = 0;
  for (; i < s.size(); ++i) {
    if (!strings::is_ascii_digit(s[i])) {
      return false;
    }
    row = row * 10 + static_cast<std::uint32_t>(s[i] - '0');
  }
  return col <= a1::kMaxCols && row >= 1 && row <= a1::kMaxRows;
}

/// Whole-string R1C1 reference: R\d*C\d*, R\d+ or C\d+ (case-insensitive).
bool is_r1c1_ref(std::string_view s) {
  if (s.empty()) {
    return false;
  }
  std::size_t i = 1;
  const char first = strings::ascii_to_upper(s[0]);
  if (first != 'R' && first != 'C') {
    return false;
  }
  std::size_t d1 = i;
  while (i < s.size() && strings::is_ascii_digit(s[i])) {
    ++i;
  }
  if (i == s.size()) {
    return i > d1;  // R\d+ or C\d+
  }
  if (first != 'R' || strings::ascii_to_upper(s[i]) != 'C') {
    return false;
  }
  ++i;
  while (i < s.size() && strings::is_ascii_digit(s[i])) {
    ++i;
  }
  return i == s.size();
}

}  // namespace

std::string copy_sheet_default_name(std::string_view source_name, const std::vector<std::string_view>& existing_names) {
  const std::string_view base = strip_copy_suffix(source_name);
  for (std::size_t n = 2;; ++n) {
    const std::string suffix = " (" + std::to_string(n) + ")";
    std::string candidate(base);
    candidate += suffix;
    if (utf16_units_in(candidate) > sheet_names::kMaxNameUnits) {
      const std::uint32_t room = sheet_names::kMaxNameUnits - utf16_units_in(suffix);
      std::string_view cut = base.substr(0, utf16_to_byte_offset_floor(base, room));
      std::string_view trimmed = cut;
      while (!trimmed.empty() && (trimmed.back() == ' ' || trimmed.back() == '\'')) {
        trimmed.remove_suffix(1);
      }
      if (!trimmed.empty()) {
        cut = trimmed;
      }
      candidate.assign(cut);
      candidate += suffix;
    }
    bool taken = false;
    for (const std::string_view existing : existing_names) {
      if (sheet_names::equal(existing, candidate)) {
        taken = true;
        break;
      }
    }
    if (!taken) {
      return candidate;
    }
  }
}

std::string copy_table_name(std::string_view source_table_name, std::size_t tables_in_workbook,
                            bool (*name_taken)(void* ctx, std::string_view candidate), void* ctx) {
  for (std::size_t n = tables_in_workbook + 2;; ++n) {
    const std::string digits = std::to_string(n);
    std::string candidate(source_table_name);
    candidate += digits;
    if (is_a1_ref(candidate) || is_r1c1_ref(candidate)) {
      candidate.assign(source_table_name);
      candidate += '_';
      candidate += digits;
    }
    if (!name_taken(ctx, candidate)) {
      return candidate;
    }
  }
}

}  // namespace formulon
