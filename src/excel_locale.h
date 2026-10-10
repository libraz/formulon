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
/// single-byte katakana; `kMacThai` and `kMacCyrillic` are the Mac pages.
enum class SbcsCodepage : std::uint8_t {
  kMacRoman = 0,
  kWindows1252 = 1,
  kCp932 = 2,
  kDbcsHighBlank = 3,
  kMacThai = 4,
  kMacCyrillic = 5,
};

/// Double-byte character set behind the byte-counting text functions and
/// CHAR / CODE above 0xFF. `kNone` is a single-byte locale.
enum class DbcsCodepage : std::uint8_t {
  kNone = 0,
  kJis0208 = 1,
  kGb2312 = 2,
  kKsX1001 = 3,
  kBig5 = 4,
};

// Criteria, lookup, and D-function text matching use the host's width and
// kana folding rules. The Windows path retains the engine's existing
// non-Mac behavior; `kNone` compares text without any width or kana fold.
enum class WidthFolding : std::uint8_t {
  kWin = 0,
  kMac = 1,
  kNone = 2,
};

/// What an `r` run in a TEXT format reads as.
enum class RLetter : std::uint8_t {
  kLiteral = 0,
  /// `r` is the era year (`ee`); a longer run adds the era name (`gggee`).
  kJapaneseEra = 1,
  /// Any run is the four-digit year.
  kYear = 2,
};

/// Letters of the localized TEXT format dialect.
struct FormatLetters {
  char year, month, day, hour, minute, second;
  /// `M` (month) and `m` (minute) are distinct letters.
  bool case_sensitive;
  /// The minute letter is a minute regardless of its neighbours.
  bool minute_unconditional;
  /// A case-sensitive month letter still reads as a minute beside an hour or second.
  bool month_contextual;
  /// Letter of the aaa / aaaa weekday tokens; `'\0'` when the locale has none.
  char weekday;
};

/// Letters used by the invariant stored format syntax and by locales whose
/// date format letters are the standard ASCII `y`, `m`, `d`, `h`, `m`, `s`.
inline constexpr FormatLetters kInvariantFormatLetters{'y', 'm', 'd', 'h', 'm', 's', false, false, false, 'a'};

/// A non-ASCII spelling of a format letter (`Д` for the day letter).
struct FormatLetterAlias {
  std::string_view spelling;
  char letter;
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
  /// A minus sign sits between a prefix symbol and the number (`¥-1,235`).
  bool minus_after_symbol;
  /// A value that display-rounds to zero keeps its negative form.
  bool negative_zero_signed;
  std::uint8_t default_decimals;
};

inline constexpr std::size_t kErrorNameCount = std::size(kErrorTable);

/// How one `[DBNumN]` directive writes numbers. Digit-by-digit output takes
/// `digits`; General, elapsed time and date fields are written with place units.
struct DbnumStyle {
  /// Digits 0-9.
  std::array<std::string_view, 10> digits;
  /// Units of 10, 100 and 1000 within a group of four digits; empty for a
  /// directive the locale accepts and leaves in ASCII digits.
  std::array<std::string_view, 3> place_units;
  /// Units of 10^4, 10^8 and 10^12.
  std::array<std::string_view, 3> group_units;
  /// A 1 before a place unit is written (一十 rather than 十).
  bool place_one;
  /// A run of skipped places between digits is written as one zero (一百○一).
  bool zero_filler;
  /// Date and time fields are written with place units rather than digit by digit.
  bool date_positional;
  /// A date field writes the 1 before the tens unit.
  bool date_place_one;
};

using DbnumStyles = std::array<DbnumStyle, 4>;

/// Names and digit styles a `[$-LCID]` format tag selects: the language's own
/// spellings, which differ from the untagged names of a profile in that locale.
struct TagLanguage {
  std::array<std::string_view, 12> month_long;
  std::array<std::string_view, 12> month_short;
  /// dddd / ddd (and aaaa / aaa) names, Sunday first.
  std::array<std::string_view, 7> day_long;
  std::array<std::string_view, 7> day_short;
  /// Written for every AM/PM, am/pm, A/P and a/p marker.
  std::string_view am_name;
  std::string_view pm_name;
  /// `[DBNum1]`-`[DBNum4]` styles; null where they change nothing.
  const DbnumStyles* dbnum;
  /// `e`, `g` and `r` write the Japanese era.
  bool japanese_era;
};

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
  /// Spellings TEXT formats write the letters of `format_letters` in; where any exist, the Latin date letters
  /// are plain text. Empty slots are unused.
  std::array<FormatLetterAlias, 10> format_letter_aliases;
  Currency currency;
  /// Currency prefixes VALUE accepts; empty slots are unused.
  std::array<std::string_view, 4> accepted_currency;
  /// Symbol USDOLLAR formats US dollars with; empty when USDOLLAR is DOLLAR.
  std::string_view usdollar_symbol;
  DateOrder date_order;
  bool dotted_date;
  bool kanji_ymd_text;
  bool kanji_time_text;
  bool japanese_era;
  RLetter r_letter;
  /// Date text takes the `2024년 3월 15일` form.
  bool hangul_ymd_text;
  /// Date text takes English month names.
  bool english_month_names;
  /// The hyphenated `d-mmm-yy` form takes the English abbreviations.
  bool hyphen_english_months;
  /// Time text takes a fraction of a second after the decimal separator.
  bool fractional_seconds;
  /// A `.` may directly follow the AM / PM marker (`6 PM.`).
  bool meridiem_dot_attached;
  /// A `.` may follow the AM / PM marker after a space (`6 PM .`).
  bool meridiem_dot_spaced;
  /// AM / PM marker spellings, which time text also accepts.
  std::string_view am_name;
  std::string_view pm_name;
  /// `A/P` is the short AM / PM marker rather than letters of the format dialect.
  bool short_meridiem;
  DbcsCodepage dbcs_codepage;
  bool halfwidth_kana_single_byte;
  bool fullwidth_numeric_text;
  bool char_snaps_near_integer;
  bool phonetic;
  bool criteria_header_keeps_halfwidth_kana;
  bool bang_escape;
  /// `[DBNum1]`-`[DBNum4]` styles; null where the directives change nothing.
  const DbnumStyles* dbnum;
  /// Under a `[$-…]` language, `[DBNum4]` keeps the profile's style in every section rather than
  /// only in the tagged one.
  bool dbnum4_keeps_profile_style;
  /// A lower-case `t` renders nothing and writes the section's digits in Thai.
  bool thai_digit_letter;
  /// Letter, in either case, of a date code that renders nothing; `'\0'` when none.
  char blank_date_letter;
  /// Full-width punctuation that renders as a literal (`／`, `％`) keeps its glyph instead of folding to ASCII.
  bool fullwidth_literal_glyph;
  /// TEXT accepts the English `General` keyword.
  bool english_general;
  /// TEXT rejects a format holding a `.` outside quotes and escapes.
  bool format_rejects_dot;
  /// Locale spelling of the General format; empty when none is measured.
  std::string_view general_alias;
  /// Colour numbers 1-8 (black blue cyan green magenta red white yellow);
  /// an empty slot is a colour name the locale rejects.
  std::array<std::string_view, 8> color_names;
  /// Spelling of the indexed colour form's prefix (`[Color12]`); empty when the locale has none.
  std::string_view color_index_prefix;
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

/// True when the localized spelling of the error at `error_ordinal`
/// (`kErrorTable` order) was captured from Excel rather than assumed.
bool error_name_measured(std::size_t error_ordinal) noexcept;
SbcsCodepage sbcs_codepage(ExcelProfile profile) noexcept;
WidthFolding width_folding(ExcelProfile profile) noexcept;

/// Language a numeric `[$-LCID]` tag names, or null when its spellings are not modelled.
const TagLanguage* tag_language_for_lcid(std::uint16_t lcid) noexcept;
/// Language a named tag (`[$-ja-JP]`, case-insensitive) names, or null when not modelled.
const TagLanguage* tag_language_for_name(std::string_view name) noexcept;
/// The tag language of the locale's own LCID.
const TagLanguage& native_tag_language(ExcelLocale locale) noexcept;
/// Digit style of the `NN` byte of `[$-NN000000]`; null for ASCII digits.
const DbnumStyle* numeral_system_style(std::uint8_t code) noexcept;
/// Invariant format codes `[$-F800]` / `[$-x-sysdate]` and `[$-F400]` / `[$-x-systime]` render through.
std::string_view system_date_format(ExcelProfile profile) noexcept;
std::string_view system_time_format(ExcelProfile profile) noexcept;

}  // namespace formulon

#endif  // FORMULON_EXCEL_LOCALE_H_
