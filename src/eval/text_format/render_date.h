// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Internal header -- do not include outside `src/eval/text_format/`.
//
// Date / time rendering for the Excel TEXT() engine. Covers the calendar
// tokens (`y`, `m`, `d`, weekday, ja-JP era), the clock tokens (`h`, `m`,
// `s`, elapsed brackets, fractional seconds), and the AM/PM / A/P
// indicators.

#ifndef FORMULON_EVAL_TEXT_FORMAT_RENDER_DATE_H_
#define FORMULON_EVAL_TEXT_FORMAT_RENDER_DATE_H_

#include <string>
#include <string_view>

#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {

// Render the date/time section for the given serial.
void render_date(const Section& section, std::string_view fmt, double serial, std::string& out);

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_RENDER_DATE_H_
