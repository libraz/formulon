// Number, boolean and error text in the active evaluation profile's locale.
// Every surface that renders these values as text goes through here so the
// locale facts are read in one place; `utils/double_format` stays invariant.

#ifndef FORMULON_EVAL_LOCALE_TEXT_H_
#define FORMULON_EVAL_LOCALE_TEXT_H_

#include <string>
#include <string_view>

#include "value.h"

namespace formulon {
namespace eval {

/// Shortest round-trip text of `value` with the locale decimal separator.
std::string locale_number_text(double value);

/// Locale name of a boolean (`TRUE` / `WAHR` / `VRAI`).
std::string_view locale_bool_text(bool value) noexcept;

/// Locale name of an error value (`#N/A` / `#NV`).
std::string_view locale_error_text(ErrorCode code) noexcept;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LOCALE_TEXT_H_
