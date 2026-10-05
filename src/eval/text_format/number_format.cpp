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

#include "eval/text_format/number_format_scanner.h"
#include "eval/text_format/number_format_types.h"
#include "eval/text_format/render_date.h"
#include "eval/text_format/render_numeric.h"
#include "utils/date_time.h"

namespace formulon {
namespace text_format {
namespace {

using number_format_detail::CondOp;
using number_format_detail::Section;

// Last serial of the 1900 calendar (9999-12-31). The 1904 system ends the
// same day, `kDate1904EpochGap` serials earlier.
constexpr double kMaxDateSerial1900 = 2958465.0;

// A format code split into sections. The views in `raw` point into
// `normalized`, so a `ParsedFormat` is filled in place and never moved.
struct ParsedFormat {
  std::string normalized;
  std::vector<std::string_view> raw;
  std::vector<Section> sections;
};

void parse_format(std::string_view format, FormatDialect dialect, ParsedFormat& parsed) {
  // Normalise the ja-JP full-width syntax once. All section/string views and
  // literal offsets below refer to this owned buffer for the duration of the
  // render; quoted and escaped payloads remain byte-for-byte unchanged.
  parsed.normalized = number_format_detail::normalize_ja_jp_format_syntax(format);
  parsed.raw = number_format_detail::split_sections(parsed.normalized);
  parsed.sections.reserve(parsed.raw.size());
  for (const auto& raw : parsed.raw) {
    Section s;
    number_format_detail::tokenize_section(raw, s, dialect);
    number_format_detail::classify(s, raw);
    parsed.sections.push_back(std::move(s));
  }
}

}  // namespace

FormatStatus apply_format(double value, std::string_view format, std::string& out, bool date1904,
                          FormatDialect dialect) {
  if (format.empty()) {
    return FormatStatus::kOk;
  }
  ParsedFormat parsed;
  parse_format(format, dialect, parsed);
  const std::vector<std::string_view>& sections_raw = parsed.raw;
  const std::vector<Section>& sections = parsed.sections;
  if (sections.empty()) {
    return FormatStatus::kOk;
  }

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
  if (any_predicate) {
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
  } else if (value < 0.0) {
    // Conditional sections do not imply a sign class, but an explicit
    // literal minus in the selected section is itself the sign glyph. Feed
    // it the magnitude so the numeric renderer does not prepend another
    // minus (e.g. `[<=0]-0.00` must render `-1.50`, not `--1.50`).
    bool has_explicit_minus = false;
    for (const number_format_detail::Token& token : section.tokens) {
      if (token.kind != number_format_detail::Tok::Literal || token.lit_end <= token.lit_begin) {
        continue;
      }
      const std::string_view literal = raw_fmt.substr(token.lit_begin, token.lit_end - token.lit_begin);
      if (literal.find('-') != std::string_view::npos) {
        has_explicit_minus = true;
        break;
      }
    }
    if (has_explicit_minus) {
      render_value = std::fabs(value);
    }
  }

  if (section.is_text) {
    number_format_detail::render_text_section(section, raw_fmt, std::string_view{}, out);
    return FormatStatus::kOk;
  }
  if (section.is_date) {
    // The calendar runs from serial 0 to 9999-12-31. The 1904 system also
    // shows a negative serial, as its magnitude behind a leading minus.
    const double max_serial = date1904 ? kMaxDateSerial1900 - date_time::kDate1904EpochGap : kMaxDateSerial1900;
    const double min_serial = date1904 ? -max_serial : 0.0;
    if (render_value < min_serial || render_value > max_serial) {
      return FormatStatus::kOverflow;
    }
    if (render_value < 0.0) {
      out.push_back('-');
      render_value = -render_value;
    }
    number_format_detail::render_date(section, raw_fmt, render_value, out, date1904);
    return FormatStatus::kOk;
  }
  number_format_detail::render_numeric(section, raw_fmt, render_value, out);
  return FormatStatus::kOk;
}

FormatStatus apply_text_format(std::string_view text, std::string_view format, std::string& out,
                               FormatDialect dialect) {
  ParsedFormat parsed;
  parse_format(format, dialect, parsed);
  std::size_t chosen = 0;
  if (parsed.sections.size() >= 4U) {
    chosen = 3U;
  } else if (parsed.sections.size() != 1U || !parsed.sections[0].is_text) {
    out.append(text);
    return FormatStatus::kOk;
  }
  const Section& section = parsed.sections[chosen];
  if (section.has_invalid_bracket) {
    return FormatStatus::kValueError;
  }
  number_format_detail::render_text_section(section, parsed.raw[chosen], text, out);
  return FormatStatus::kOk;
}

bool negative_section_has_color(std::string_view format, FormatDialect dialect) {
  ParsedFormat parsed;
  parse_format(format, dialect, parsed);
  if (parsed.sections.empty()) {
    return false;
  }
  return parsed.sections[parsed.sections.size() >= 2U ? 1U : 0U].has_color;
}

}  // namespace text_format
}  // namespace formulon
