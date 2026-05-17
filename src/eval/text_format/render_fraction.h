// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Internal header -- do not include outside `src/eval/text_format/`.
//
// Fraction rendering for the Excel TEXT() engine (`# ?/?`, `# ??/??`,
// `0/0`, ...). Implements Excel's bounded best-rational-approximation
// via a Stern-Brocot mediant search.

#ifndef FORMULON_EVAL_TEXT_FORMAT_RENDER_FRACTION_H_
#define FORMULON_EVAL_TEXT_FORMAT_RENDER_FRACTION_H_

#include <string>
#include <string_view>

#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {

// Render one fraction-format section. The dispatcher in `render_numeric`
// calls this when `Section::is_fraction` is true.
void render_fraction(const Section& section, std::string_view fmt, double value, std::string& out);

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_RENDER_FRACTION_H_
