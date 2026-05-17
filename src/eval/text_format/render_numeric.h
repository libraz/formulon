// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Internal header -- do not include outside `src/eval/text_format/`.
//
// Standard numeric / `General` rendering for the Excel TEXT() engine.
// See `render_date.h` for date/time and `render_fraction.h` for fractions.

#ifndef FORMULON_EVAL_TEXT_FORMAT_RENDER_NUMERIC_H_
#define FORMULON_EVAL_TEXT_FORMAT_RENDER_NUMERIC_H_

#include <string>
#include <string_view>

#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {

// Render one numeric section through the walk-tokens pipeline. Delegates to
// `render_fraction` when `section.is_fraction` is set.
void render_numeric(const Section& section, std::string_view fmt, double value, std::string& out);

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_RENDER_NUMERIC_H_
