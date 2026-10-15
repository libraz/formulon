//
// The date / time text recogniser shared by VALUE, NUMBERVALUE, DATEVALUE,
// TIMEVALUE, implicit text->number coercion and the criteria engine. The
// pre-processing, the time grammar and the date+time composition rules are
// the same in every profile; the active profile supplies the date shapes:
// Year-Month-Day profiles read numeric pairs and triples, English month
// names and the kanji / era / hangul forms as measured on Mac Excel ja-JP,
// the others keep their locale-ordered numeric and month-name forms.

#ifndef FORMULON_EVAL_DATE_TEXT_PARSE_H_
#define FORMULON_EVAL_DATE_TEXT_PARSE_H_

#include <cstdint>
#include <string_view>

namespace formulon {
namespace eval {
namespace date_parse {

/// How U+0020 blanks at the ends of the text are treated.
enum class TextEdges : std::uint8_t {
  kTrim,      ///< VALUE, DATEVALUE, TIMEVALUE and criteria strip them at both ends.
  kKeepLead,  ///< Implicit coercion strips only trailing blanks; a leading one rejects the text.
};

/// A recognised date / time text.
struct DateTimeText {
  /// Date serial (1900 system) plus the time of day; hours past 24 and the
  /// millisecond rounding carry into the integer part.
  double value;
  /// The time part alone (0 without one), hours past 24 included.
  double time;
  /// True when the text held a date; a time-only text counts days from 0.
  bool has_date;
};

/// Recognises `text` as a date, a time, or a date and a time in either
/// order, and writes the result into `*out`. Texts longer than 255 UTF-16
/// units, and texts holding control characters or blanks other than U+0020 /
/// U+3000, are rejected. `current_year` supplies the year of the year-less
/// forms (`1/2`, `Jan 1`, `1月2日`); `0` disables them. Returns false, leaving
/// `*out` untouched, when the text is not a date or time.
bool parse_date_time_text(std::string_view text, TextEdges edges, int current_year, DateTimeText* out) noexcept;

}  // namespace date_parse
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DATE_TEXT_PARSE_H_
