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
  using number_format_detail::CondOp;
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
    number_format_detail::classify(s, raw);
    sections.push_back(std::move(s));
  }

  // Caller is passing `original_text`: if the value is text (non-numeric
  // source) and we have a text section (index 3 for 4-section formats; any
  // `@` token in a single-section format also applies), route there.
  const bool has_original_text = !original_text.empty();

  // Returns true if `op(v, pred)` holds; `kNone` is treated as the always-true
  // unconditional sentinel.
  auto cond_match = [](CondOp op, double pred, double v) -> bool {
    switch (op) {
      case CondOp::kGt:
        return v > pred;
      case CondOp::kGe:
        return v >= pred;
      case CondOp::kLt:
        return v < pred;
      case CondOp::kLe:
        return v <= pred;
      case CondOp::kEq:
        return v == pred;
      case CondOp::kNe:
        return v != pred;
      case CondOp::kNone:
        return true;
    }
    return false;
  };

  // Predicate-based dispatch (`[>N]` / `[<N]` / `[=N]` / ... section prefix).
  // Excel's rule: when section 0 OR section 1 carries a predicate, sign-class
  // dispatch is replaced by predicate matching. Sections are visited in
  // declaration order; the first whose predicate holds wins. With three or
  // more sections, section 2 is the unconditional fallback when neither
  // predicate matches.
  bool used_conditional = false;
  // Decide the section to use based on Excel's rules:
  //   1 section : apply to everything; text passes unformatted unless `@`
  //               is present.
  //   2 sections: section 0 = positive/zero; section 1 = negative.
  //   3 sections: section 0 = positive; section 1 = negative; section 2 = zero.
  //   4 sections: section 0 = positive; section 1 = negative; section 2 = zero;
  //               section 3 = text.
  int chosen = 0;
  const bool any_predicate = (!sections.empty() && sections[0].cond_op != CondOp::kNone) ||
                             (sections.size() >= 2 && sections[1].cond_op != CondOp::kNone);
  if (has_original_text) {
    if (sections.size() >= 4) {
      chosen = 3;
    } else {
      // Single-section with an `@`: route through the numeric walker but
      // `@` substitutes the text.
      chosen = 0;
    }
  } else if (any_predicate) {
    used_conditional = true;
    chosen = -1;
    // Walk the first two sections, picking the first whose predicate holds.
    // A section without a predicate (`cond_op == kNone`) acts as the
    // catch-all in this position.
    const std::size_t scan_limit = sections.size() < 2 ? sections.size() : 2;
    for (std::size_t i = 0; i < scan_limit; ++i) {
      if (cond_match(sections[i].cond_op, sections[i].cond_value, value)) {
        chosen = static_cast<int>(i);
        break;
      }
    }
    if (chosen < 0) {
      // Neither of sections 0/1 matched. Excel uses section 2 as the
      // unconditional fallback when present; otherwise it falls through to
      // section 1 (the predicateless section, by elimination) so that the
      // user-supplied "else" arm renders. If both arms had predicates, fall
      // back to section 0 to mirror Mac Excel's "first section wins" tiebreak.
      if (sections.size() >= 3) {
        chosen = 2;
      } else if (sections.size() >= 2 && sections[1].cond_op == CondOp::kNone) {
        chosen = 1;
      } else if (!sections.empty() && sections[0].cond_op == CondOp::kNone) {
        chosen = 0;
      } else {
        chosen = 0;
      }
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
  //
  // When predicate-based dispatch picked the section, the chosen index no
  // longer correlates with sign class — the value's sign should be rendered
  // verbatim. Skip the abs-adjustment in that case.
  double render_value = value;
  if (!used_conditional) {
    if (chosen == 1 && sections.size() >= 2) {
      render_value = std::fabs(value);
    } else if (chosen == 2 && sections.size() >= 3) {
      render_value = std::fabs(value);
    }
  }

  if (section.is_text ||
      (has_original_text && !section.is_date && section.integer_zero_digits == 0 && section.integer_opt_digits == 0 &&
       section.integer_pad_digits == 0 && section.fraction_zero_digits == 0 && section.fraction_opt_digits == 0 &&
       section.fraction_pad_digits == 0)) {
    number_format_detail::render_text_section(section, raw_fmt, original_text, out);
    return FormatStatus::kOk;
  }
  if (section.is_date) {
    // Excel rejects out-of-range date serials from TEXT: the valid range is
    // [0, 2958465] (the latter is 9999-12-31). Surface as #VALUE! rather than
    // silently emitting an empty string.
    if (render_value < 0.0 || render_value > 2958465.0) {
      return FormatStatus::kValueError;
    }
    number_format_detail::render_date(section, raw_fmt, render_value, out);
    return FormatStatus::kOk;
  }
  number_format_detail::render_numeric(section, raw_fmt, render_value, out);
  return FormatStatus::kOk;
}

}  // namespace text_format
}  // namespace eval
}  // namespace formulon
