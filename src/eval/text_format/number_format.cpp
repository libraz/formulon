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

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "eval/text_format/format_localize.h"
#include "eval/text_format/number_format_types.h"
#include "eval/text_format/render_date.h"
#include "eval/text_format/render_numeric.h"
#include "excel_locale.h"
#include "utils/date_time.h"
#include "utils/subnormal_flush.h"

namespace formulon {
namespace text_format {
namespace {

using number_format_detail::CondOp;
using number_format_detail::Section;

// A format code split into sections. The views in `raw` point into
// `normalized`, so a `ParsedFormat` is filled in place and never moved.
struct ParsedFormat {
  std::string normalized;
  bool spelling_valid = true;
  std::vector<std::string_view> raw;
  std::vector<Section> sections;
};

void parse_format(std::string_view format, FormatDialect dialect, ParsedFormat& parsed) {
  const ExcelProfile profile = eval::current_eval_profile();
  // All section/string views and literal offsets below refer to this owned
  // invariant-syntax buffer for the duration of the render.
  number_format_detail::LocalizedFormat localized = number_format_detail::localize_format(format, dialect, profile);
  parsed.normalized = std::move(localized.text);
  parsed.spelling_valid = localized.valid;
  parsed.raw = number_format_detail::split_sections(parsed.normalized);
  // A TEXT() argument reads dates in the locale's letters; a stored code in the invariant ones.
  const FormatLetters& letters = dialect == FormatDialect::kLocalized ? locale_facts(profile).format_letters
                                                                      : number_format_detail::kInvariantFormatLetters;
  parsed.sections.reserve(parsed.raw.size());
  for (const auto& raw : parsed.raw) {
    Section s;
    number_format_detail::tokenize_section(raw, s, letters,
                                           parsed.sections.empty() ? nullptr : &parsed.sections.front().tag);
    number_format_detail::classify(s, raw);
    parsed.sections.push_back(std::move(s));
  }
}

// The parse of the last format this thread asked for, keyed on everything `parse_format` reads: the text, the dialect
// and the evaluation profile. A TEXT() over a range repeats one format per cell. The handle keeps a parse alive
// across a nested call that replaces the entry.
std::shared_ptr<const ParsedFormat> parse_format_cached(std::string_view format, FormatDialect dialect) {
  struct Entry {
    std::string format;
    FormatDialect dialect = FormatDialect::kLocalized;
    ExcelProfile profile;
    std::shared_ptr<const ParsedFormat> parsed;
  };
  thread_local Entry entry;
  const ExcelProfile profile = eval::current_eval_profile();
  if (entry.parsed != nullptr && entry.dialect == dialect && entry.profile.host == profile.host &&
      entry.profile.locale == profile.locale && entry.format == format) {
    return entry.parsed;
  }
  auto parsed = std::make_shared<ParsedFormat>();
  parse_format(format, dialect, *parsed);
  entry.format.assign(format);
  entry.dialect = dialect;
  entry.profile = profile;
  entry.parsed = std::move(parsed);
  return entry.parsed;
}

// A number format asks for at most this many decimal places in a TEXT() argument.
constexpr int kMaxFormatDecimalPlaces = 127;

// The most digit placeholders (`0`, `#`, `?`) following a decimal point in any section of a stored-dialect `format`;
// quoted text, escaped characters and bracketed parts are skipped.
int format_decimal_places(std::string_view format) {
  int most = 0;
  int run = 0;
  bool after_point = false;
  for (std::size_t i = 0; i < format.size(); ++i) {
    const char c = format[i];
    if (c == '"') {
      const std::size_t close = format.find('"', i + 1);
      i = close == std::string_view::npos ? format.size() : close;
    } else if (c == '\\') {
      ++i;
    } else if (c == '[') {
      const std::size_t close = format.find(']', i + 1);
      i = close == std::string_view::npos ? format.size() : close;
    } else if (c == ';') {
      after_point = false;
      run = 0;
    } else if (c == '.') {
      after_point = true;
      run = 0;
    } else if (after_point && (c == '0' || c == '#' || c == '?')) {
      most = std::max(most, ++run);
    } else {
      after_point = false;
      run = 0;
    }
  }
  return most;
}

bool valid_format(const ParsedFormat& parsed) noexcept {
  if (!parsed.spelling_valid || parsed.sections.size() > 4U) {
    return false;
  }
  for (std::size_t i = 0; i < parsed.sections.size(); ++i) {
    const Section& section = parsed.sections[i];
    if (section.has_invalid_bracket || (i >= 2U && section.cond_op != CondOp::kNone)) {
      return false;
    }
  }
  return true;
}

// Returns true if `op(v, pred)` holds; `kNone` is the always-true
// unconditional sentinel.
bool cond_match(CondOp op, double pred, double v) {
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
}

// The section a value renders with. `conditional` is set when a `[>N]` /
// `[<N]` / ... predicate on section 0 or 1 decided it (predicate matching then
// replaces sign-class dispatch); `predicate_matched` says the chosen arm's own
// predicate held.
struct SectionChoice {
  int chosen = 0;
  bool predicate_matched = false;
  bool conditional = false;
};

// Decides the section by Excel's rules:
//   1 section : apply to everything; text passes unformatted unless `@`
//               is present.
//   2 sections: section 0 = positive/zero; section 1 = negative.
//   3 sections: section 0 = positive; section 1 = negative; section 2 = zero.
//   4 sections: section 0 = positive; section 1 = negative; section 2 = zero;
//               section 3 = text.
// Returns false when a two-section predicate format has no arm for `value`.
// Each predicate branch states what an unconditional arm stands for.
bool choose_section(const std::vector<Section>& sections, double value, SectionChoice* out) {
  SectionChoice choice;
  const bool any_predicate =
      sections[0].cond_op != CondOp::kNone || (sections.size() >= 2 && sections[1].cond_op != CondOp::kNone);
  if (any_predicate) {
    choice.conditional = true;
    const bool first_predicate = sections[0].cond_op != CondOp::kNone;
    const bool second_predicate = sections.size() >= 2U && sections[1].cond_op != CondOp::kNone;
    if (sections.size() >= 3U) {
      // Unconditional arms 0/1 are implicit positive/negative predicates; section 2 is the raw fallback.
      const bool first_matches =
          first_predicate ? cond_match(sections[0].cond_op, sections[0].cond_value, value) : value > 0.0;
      if (first_matches) {
        choice.chosen = 0;
        choice.predicate_matched = first_predicate;
      } else {
        const bool second_matches =
            second_predicate ? cond_match(sections[1].cond_op, sections[1].cond_value, value) : value < 0.0;
        if (second_matches) {
          choice.chosen = 1;
          choice.predicate_matched = second_predicate;
        } else {
          choice.chosen = 2;
        }
      }
    } else if (sections.size() == 2U) {
      if (first_predicate && second_predicate) {
        if (cond_match(sections[0].cond_op, sections[0].cond_value, value)) {
          choice.chosen = 0;
          choice.predicate_matched = true;
        } else if (cond_match(sections[1].cond_op, sections[1].cond_value, value)) {
          choice.chosen = 1;
          choice.predicate_matched = true;
        } else {
          return false;
        }
      } else if (first_predicate) {
        if (cond_match(sections[0].cond_op, sections[0].cond_value, value)) {
          choice.chosen = 0;
          choice.predicate_matched = true;
        } else {
          // The unconditional second arm is the fallback when the first arm carries the predicate.
          choice.chosen = 1;
        }
      } else {
        // An unconditional first arm is the implicit positive arm; zero belongs to neither arm.
        if (value > 0.0) {
          choice.chosen = 0;
        } else if (cond_match(sections[1].cond_op, sections[1].cond_value, value)) {
          choice.chosen = 1;
          choice.predicate_matched = true;
        } else {
          return false;
        }
      }
    } else {
      // A lone predicate section still renders on a miss; the sign treatment tells the cases apart.
      choice.chosen = 0;
      choice.predicate_matched = cond_match(sections[0].cond_op, sections[0].cond_value, value);
    }
  } else if (value > 0.0) {
    choice.chosen = 0;
  } else if (value < 0.0) {
    choice.chosen = sections.size() >= 2 ? 1 : 0;
  } else {
    choice.chosen = sections.size() >= 3 ? 2 : 0;
  }
  *out = choice;
  return true;
}

// Whether the chosen section renders the magnitude of `value` rather than the
// signed value. For section 1 (negative) Excel renders the magnitude and any
// explicit sign comes from that section's literals; the numeric walker
// prefixes a minus for negative input. Predicate dispatch chooses from the
// selected explicit or implicit arm, since section index alone does not
// determine it.
bool renders_magnitude(const std::vector<Section>& sections, const SectionChoice& choice, double value) {
  const Section& section = sections[static_cast<std::size_t>(choice.chosen)];
  if (!choice.conditional) {
    return choice.chosen == 1 || choice.chosen == 2;
  }
  if (!(value < 0.0)) {
    return false;
  }
  if (section.cond_op != CondOp::kNone) {
    if (choice.predicate_matched) {
      // Excel passes the magnitude for `<` at a threshold <= 0, and for `<=`/`=` at a threshold < 0.
      switch (section.cond_op) {
        case CondOp::kLt:
          return section.cond_value <= 0.0;
        case CondOp::kLe:
        case CondOp::kEq:
          return section.cond_value < 0.0;
        default:
          return false;
      }
    }
    if (sections.size() == 1U) {
      // A negative miss renders positively for negative-only or all-positive predicates; `=` stays signed.
      switch (section.cond_op) {
        case CondOp::kLt:
          return section.cond_value <= 0.0;
        case CondOp::kLe:
          return section.cond_value < 0.0;
        case CondOp::kGt:
        case CondOp::kGe:
          return section.cond_value <= 0.0;
        case CondOp::kNe:
          return section.cond_value < 0.0;
        case CondOp::kEq:
        case CondOp::kNone:
          return false;
      }
    }
    return false;
  }
  if (sections.size() >= 3U && choice.chosen == 1) {
    // Section 1 is the implicit negative-magnitude arm; section 2 is the raw fallback and keeps the sign.
    return true;
  }
  if (sections.size() == 2U && sections[0].cond_op != CondOp::kNone) {
    // Only a predicate covering every positive value makes the second arm the negative section.
    const CondOp first_op = sections[0].cond_op;
    return ((first_op == CondOp::kGt || first_op == CondOp::kGe) && sections[0].cond_value <= 0.0) ||
           (first_op == CondOp::kNe && sections[0].cond_value < 0.0);
  }
  return false;
}

}  // namespace

FormatStatus apply_format(double value, std::string_view format, std::string& out, bool date1904, FormatDialect dialect,
                          bool cap_decimals) {
  if (format.empty()) {
    return FormatStatus::kOk;
  }
  // Section choice and predicates use Excel's one-sided zero test.
  if (tests_as_zero(value)) {
    value = 0.0;
  }
  const std::shared_ptr<const ParsedFormat> held = parse_format_cached(format, dialect);
  const ParsedFormat& parsed = *held;
  if (!valid_format(parsed) || (cap_decimals && format_decimal_places(parsed.normalized) > kMaxFormatDecimalPlaces)) {
    return FormatStatus::kValueError;
  }
  const std::vector<std::string_view>& sections_raw = parsed.raw;
  const std::vector<Section>& sections = parsed.sections;
  if (sections.empty()) {
    return FormatStatus::kOk;
  }

  SectionChoice choice;
  if (!choose_section(sections, value, &choice)) {
    return FormatStatus::kOverflow;
  }
  const int chosen = choice.chosen;
  const Section& section = sections[static_cast<std::size_t>(chosen)];
  const std::string_view raw_fmt = sections_raw[static_cast<std::size_t>(chosen)];

  double render_value = renders_magnitude(sections, choice, value) ? std::fabs(value) : value;
  if (section.tag.system != number_format_detail::SystemFormat::kNone) {
    // `[$-F800]` / `[$-F400]` show the value in the system's own date or time
    // form, whatever else the section holds (locale_tokens.text_system_date_tag).
    if (value < 0.0) {
      return FormatStatus::kValueError;
    }
    const ExcelProfile profile = eval::current_eval_profile();
    const std::string_view system = section.tag.system == number_format_detail::SystemFormat::kLongDate
                                        ? system_date_format(profile)
                                        : system_time_format(profile);
    return apply_format(value, system, out, date1904, FormatDialect::kStored);
  }
  if (section.is_text) {
    // A number selecting a text-only section (`@`, `"pre"@`) renders as General without its
    // literals, in the format's numeral system (locale_tokens.lcid_section_scope).
    std::string general = "General";
    if (const std::uint8_t numeral = sections.front().tag.numeral; numeral != 0) {
      constexpr std::string_view kHex = "0123456789ABCDEF";
      general = std::string("[$-") + kHex[numeral >> 4U] + kHex[numeral & 0xFU] + "000000]General";
    }
    return apply_format(value, general, out, date1904, FormatDialect::kStored);
  }
  if (section.is_date) {
    // The calendar runs from serial 0 to 9999-12-31. The 1904 system also
    // shows a negative serial, as its magnitude behind a leading minus.
    const double max_serial =
        date1904 ? date_time::kMaxSerial1900 - date_time::kDate1904EpochGap : date_time::kMaxSerial1900;
    // The last calendar day keeps its time fraction; the next whole day is out of range in either epoch.
    const bool out_of_range =
        date1904 ? std::fabs(render_value) >= max_serial + 1.0 : render_value < 0.0 || render_value >= max_serial + 1.0;
    if (out_of_range) {
      return FormatStatus::kOverflow;
    }
    const std::size_t out_size = out.size();
    if (render_value < 0.0) {
      out.push_back('-');
      render_value = -render_value;
    }
    const FormatStatus status = number_format_detail::render_date(section, raw_fmt, render_value, out, date1904);
    if (status != FormatStatus::kOk) {
      out.resize(out_size);
    }
    return status;
  }
  return number_format_detail::render_numeric(section, raw_fmt, render_value, out);
}

FormatStatus apply_text_format(std::string_view text, std::string_view format, std::string& out,
                               FormatDialect dialect) {
  const std::shared_ptr<const ParsedFormat> held = parse_format_cached(format, dialect);
  const ParsedFormat& parsed = *held;
  // Even when the format has no text section and would otherwise pass the
  // original bytes through unchanged, Excel validates every section first.
  // This matters for text/bool TEXT values with a malformed numeric sibling
  // section (for example `@;[invalid]0`).
  if (!valid_format(parsed)) {
    return FormatStatus::kValueError;
  }
  std::size_t chosen = 0;
  if (parsed.sections.size() >= 4U) {
    chosen = 3U;
  } else if (parsed.sections.size() != 1U || !parsed.sections[0].is_text) {
    out.append(text);
    return FormatStatus::kOk;
  }
  const Section& section = parsed.sections[chosen];
  number_format_detail::render_text_section(section, parsed.raw[chosen], text, out);
  return FormatStatus::kOk;
}

bool negative_section_has_color(std::string_view format, FormatDialect dialect) {
  const std::shared_ptr<const ParsedFormat> held = parse_format_cached(format, dialect);
  const ParsedFormat& parsed = *held;
  if (!parsed.spelling_valid || parsed.sections.empty()) {
    return false;
  }
  return parsed.sections[parsed.sections.size() >= 2U ? 1U : 0U].has_color;
}

}  // namespace text_format
}  // namespace formulon
