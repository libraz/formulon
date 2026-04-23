// Copyright 2026 libraz. Licensed under the MIT License.
//
// Public entry point for the Excel TEXT() format-string engine declared in
// `number_format.h`. The design follows the two-phase approach described
// in the scope memo: (1) tokenize the format, splitting on `;` into up to
// four sections; (2) render a value through the section selected by its
// sign/zero/text classification.
//
// The tokenizer lives in `number_format_tokenizer.cpp`; rendering (numeric,
// date, and text) lives in `number_format_render.cpp`. Shared types are
// declared in the non-public `number_format_types.h`.

#include "eval/text_format/number_format.h"

#include <cmath>
#include <cstddef>
#include <string>
#include <string_view>
#include <vector>

#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace eval {
namespace text_format {

FormatStatus apply_format(double value, std::string_view format, std::string_view original_text, std::string& out) {
  using number_format_detail::Section;
  if (format.empty()) {
    return FormatStatus::kOk;
  }
  const auto sections_raw = number_format_detail::split_sections(format);
  if (sections_raw.empty()) {
    return FormatStatus::kOk;
  }
  std::vector<Section> sections;
  sections.reserve(sections_raw.size());
  for (const auto& raw : sections_raw) {
    Section s;
    number_format_detail::tokenize_section(raw, s);
    number_format_detail::classify(s);
    sections.push_back(std::move(s));
  }

  // Caller is passing `original_text`: if the value is text (non-numeric
  // source) and we have a text section (index 3 for 4-section formats; any
  // `@` token in a single-section format also applies), route there.
  const bool has_original_text = !original_text.empty();
  // Decide the section to use based on Excel's rules:
  //   1 section : apply to everything; text passes unformatted unless `@`
  //               is present.
  //   2 sections: section 0 = positive/zero; section 1 = negative.
  //   3 sections: section 0 = positive; section 1 = negative; section 2 = zero.
  //   4 sections: section 0 = positive; section 1 = negative; section 2 = zero;
  //               section 3 = text.
  int chosen = 0;
  if (has_original_text) {
    if (sections.size() >= 4) {
      chosen = 3;
    } else {
      // Single-section with an `@`: route through the numeric walker but
      // `@` substitutes the text.
      chosen = 0;
    }
  } else if (value > 0.0) {
    chosen = 0;
  } else if (value < 0.0) {
    if (sections.size() >= 2) {
      chosen = 1;
    } else {
      chosen = 0;
    }
  } else {
    // Zero.
    if (sections.size() >= 3) {
      chosen = 2;
    } else {
      chosen = 0;
    }
  }

  const Section& section = sections[static_cast<std::size_t>(chosen)];
  const std::string_view raw_fmt = sections_raw[static_cast<std::size_t>(chosen)];
  if (section.has_invalid_bracket) {
    return FormatStatus::kValueError;
  }

  // For section 1 (negative) Excel emits the value's absolute representation
  // unless the format itself includes an explicit minus sign. The numeric
  // walker currently prefixes the minus from `signbit(scaled)`, so pass the
  // absolute value when we've chosen the dedicated negative section.
  double render_value = value;
  if (chosen == 1 && sections.size() >= 2) {
    render_value = std::fabs(value);
  } else if (chosen == 2 && sections.size() >= 3) {
    render_value = std::fabs(value);
  }

  if (section.is_text ||
      (has_original_text && !section.is_date && section.integer_zero_digits == 0 && section.integer_opt_digits == 0 &&
       section.integer_pad_digits == 0 && section.fraction_zero_digits == 0 && section.fraction_opt_digits == 0 &&
       section.fraction_pad_digits == 0)) {
    number_format_detail::render_text_section(section, raw_fmt, original_text, out);
    return FormatStatus::kOk;
  }
  if (section.is_date) {
    number_format_detail::render_date(section, raw_fmt, render_value, out);
    return FormatStatus::kOk;
  }
  number_format_detail::render_numeric(section, raw_fmt, render_value, out);
  return FormatStatus::kOk;
}

}  // namespace text_format
}  // namespace eval
}  // namespace formulon
