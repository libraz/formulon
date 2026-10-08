#ifndef FORMULON_EXCEL_LOCALE_H_
#define FORMULON_EXCEL_LOCALE_H_

#include <array>
#include <cstddef>
#include <cstdint>
#include <iterator>
#include <string_view>

#include "excel_profile.h"
#include "value.h"

namespace formulon {

enum class DateOrder : std::uint8_t {
  kMDY = 0,
  kYMD = 1,
  kDMY = 2,
};

/// Single-byte code page used by CHAR / CODE for bytes 0x80-0xFF.
/// `kDbcsHighBlank` decodes the high bytes of a DBCS locale without
/// single-byte katakana; `kMacThai` is the Mac Thai page.
enum class SbcsCodepage : std::uint8_t {
  kMacRoman = 0,
  kWindows1252 = 1,
  kCp932 = 2,
  kDbcsHighBlank = 3,
  kMacThai = 4,
};

/// Double-byte character set behind the byte-counting text functions and
/// CHAR / CODE above 0xFF. `kNone` is a single-byte locale.
enum class DbcsCodepage : std::uint8_t {
  kNone = 0,
  kJis0208 = 1,
  kGb2312 = 2,
  kKsX1001 = 3,
};

// Criteria, lookup, and D-function text matching use the host's width and
// kana folding rules. The Windows path retains the engine's existing
// non-Mac behavior; `kNone` compares text without any width or kana fold.
enum class WidthFolding : std::uint8_t {
  kWin = 0,
  kMac = 1,
  kNone = 2,
};

/// Letters of the localized TEXT format dialect.
struct FormatLetters {
  char year, month, day, hour, minute, second;
  /// `M` (month) and `m` (minute) are distinct letters.
  bool case_sensitive;
  /// The minute letter is a minute regardless of its neighbours.
  bool minute_unconditional;
  /// Letter of the aaa / aaaa weekday tokens; `'\0'` when the locale has none.
  char weekday;
};

/// Locale currency used by DOLLAR and the currency-aware renderers.
struct Currency {
  std::string_view symbol;
  /// The symbol follows the number (`1,00 €`).
  bool suffix;
  /// One ASCII space separates the symbol from the number.
  bool space;
  /// Negative amounts are parenthesised rather than minus-signed.
  bool negative_parens;
  /// A value that display-rounds to zero keeps its negative form.
  bool negative_zero_signed;
  std::uint8_t default_decimals;
};

inline constexpr std::size_t kErrorNameCount = std::size(kErrorTable);

struct LocaleFacts {
  char decimal_separator;
  char group_separator;
  char list_separator;
  char array_column_separator;
  char array_row_separator;
  std::string_view true_name;
  std::string_view false_name;
  /// AND / OR / XOR / IFS skip `TRUE` / `FALSE` text like range text
  /// instead of failing on it.
  bool logical_skips_english_bool_text;
  /// Indexed by `ErrorCode` ordinal, in `kErrorTable` order.
  std::array<std::string_view, kErrorNameCount> error_names;
  char r1c1_row;
  char r1c1_col;
  char r1c1_open;
  char r1c1_close;
  char cell_type_blank;
  char cell_type_label;
  char cell_type_value;
  char cell_format_general;
  std::array<std::string_view, 12> month_long;
  std::array<std::string_view, 12> month_short;
  /// ddd / dddd names, Sunday first.
  std::array<std::string_view, 7> day_long;
  std::array<std::string_view, 7> day_short;
  FormatLetters format_letters;
  Currency currency;
  /// Currency prefixes VALUE accepts; empty slots are unused.
  std::array<std::string_view, 4> accepted_currency;
  /// USDOLLAR formats in US dollars; otherwise it is DOLLAR.
  bool usdollar_in_dollars;
  DateOrder date_order;
  bool dotted_date;
  bool kanji_ymd_text;
  bool kanji_time_text;
  bool japanese_era;
  /// Date text takes the `2024년 3월 15일` form.
  bool hangul_ymd_text;
  /// Date text takes English month names; the hyphenated `d-mmm-yy` form
  /// takes the English abbreviations everywhere.
  bool english_month_names;
  /// Time text takes a fraction of a second after the decimal separator.
  bool fractional_seconds;
  /// A `.` may directly follow the AM / PM marker (`6 PM.`).
  bool meridiem_dot_attached;
  /// A `.` may follow the AM / PM marker after a space (`6 PM .`).
  bool meridiem_dot_spaced;
  DbcsCodepage dbcs_codepage;
  bool halfwidth_kana_single_byte;
  bool fullwidth_numeric_text;
  bool char_snaps_near_integer;
  bool phonetic;
  bool criteria_header_keeps_halfwidth_kana;
  bool bang_escape;
  bool dbnum;
  bool fullwidth_syntax_fold;
  /// Locale spelling of the General format; empty when only `General` exists.
  std::string_view general_alias;
  /// Colour numbers 1-8 (black blue cyan green magenta red white yellow);
  /// an empty slot is a colour name the locale rejects.
  std::array<std::string_view, 8> color_names;
  /// [DBNum1] and [DBNum2] digits 0-9.
  std::array<std::array<std::string_view, 10>, 2> dbnum_digits;
  std::string_view grand_total;
  std::string_view hierarchy_grand_total;
  std::string_view row_field_prefix;
  std::string_view column_field_prefix;
  std::string_view value_field_prefix;
  /// aaa / aaaa names, Sunday first.
  std::array<std::string_view, 7> weekday_long;
  std::array<std::string_view, 7> weekday_short;
};

const LocaleFacts& locale_facts(ExcelProfile profile) noexcept;
SbcsCodepage sbcs_codepage(ExcelProfile profile) noexcept;
WidthFolding width_folding(ExcelProfile profile) noexcept;

}  // namespace formulon

#endif  // FORMULON_EXCEL_LOCALE_H_
