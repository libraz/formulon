//
// Shared definition of `read_int_arg` for the text builtin family. Hoisted
// out of `text.cpp` into its own TU because the DBCS family (`text_dbcs.cpp`)
// and the modern TEXTBEFORE/TEXTAFTER family (`text_modern.cpp`) both
// reference it; keeping a single definition here avoids ODR violations.

#include "eval/builtins/text_detail.h"

#include <cmath>
#include <utility>

#include "eval/coerce.h"
#include "eval/eval_profile_scope.h"
#include "eval/wildcard.h"
#include "excel_locale.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"

namespace formulon {
namespace eval {
namespace text_detail {

namespace {

Expected<double, ErrorCode> read_finite_number(const Value& v) {
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return std::move(coerced.error());
  }
  const double d = coerced.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  return d;
}

// SEARCHB's `?` spans one SBCS character in ja and ko but any character in zh
// (locale_tokens.searchb_question_kanji).
bool question_spans_sbcs_only(DbcsCodepage codepage) noexcept {
  return codepage != DbcsCodepage::kGb2312;
}

}  // namespace

int dbcs_char_bytes(std::uint32_t codepoint, bool halfwidth_kana_single_byte) noexcept {
  if (codepoint <= 0x7Fu) {
    return 1;
  }
  if (halfwidth_kana_single_byte && codepoint >= 0xFF61u && codepoint <= 0xFF9Fu) {
    return 1;
  }
  return 2;
}

std::uint64_t dbcs_bytes_in(std::string_view s, bool halfwidth_kana_single_byte) noexcept {
  std::uint64_t bytes = 0;
  std::size_t i = 0;
  while (i < s.size()) {
    std::size_t step = 0;
    const std::uint32_t cp = decode_utf8_step(s, i, &step);
    bytes +=
        (step == 1 && cp == 0xFFFDu) ? 1u : static_cast<std::uint64_t>(dbcs_char_bytes(cp, halfwidth_kana_single_byte));
    i += step;
  }
  return bytes;
}

Expected<int, ErrorCode> read_int_arg(const Value& v) {
  auto d = read_finite_number(v);
  if (!d) {
    return std::move(d.error());
  }
  return truncate_saturated_int(d.value());
}

Expected<int, ErrorCode> read_snapped_int_arg(const Value& v, double min) {
  auto d = read_finite_number(v);
  if (!d) {
    return std::move(d.error());
  }
  if (d.value() < min) {
    return ErrorCode::Value;
  }
  return truncate_saturated_int(snap_near_integer(d.value()));
}

Expected<int, ErrorCode> read_optional_int_arg(const Value* args, std::uint32_t arity, std::uint32_t index,
                                               int default_value) {
  if (arity <= index) {
    return default_value;
  }
  return read_int_arg(args[index]);
}

bool read_search_args(const Value* args, std::uint32_t arity, SearchUnit unit, SearchArgs* out, Value* out_result) {
  auto needle = coerce_to_text(args[0]);
  if (!needle) {
    *out_result = Value::error(needle.error());
    return false;
  }
  auto haystack = coerce_to_text(args[1]);
  if (!haystack) {
    *out_result = Value::error(haystack.error());
    return false;
  }
  auto parsed = arity > 2u ? read_snapped_int_arg(args[2], 1.0) : Expected<int, ErrorCode>(1);
  if (!parsed) {
    *out_result = Value::error(parsed.error());
    return false;
  }
  const int start = parsed.value();
  const std::uint64_t total =
      (unit == SearchUnit::DbcsByte)
          ? dbcs_bytes_in(haystack.value(), locale_facts(current_eval_profile()).halfwidth_kana_single_byte)
          : static_cast<std::uint64_t>(utf16_units_in(haystack.value()));
  if (start < 1 || static_cast<std::uint64_t>(start) > total + 1) {
    *out_result = Value::error(ErrorCode::Value);
    return false;
  }
  if (needle.value().empty()) {
    *out_result = Value::number(static_cast<double>(start));
    return false;
  }
  out->needle = std::move(needle.value());
  out->haystack = std::move(haystack.value());
  out->start = start;
  return true;
}

std::size_t find_folded(const std::string& haystack, const std::string& needle, std::size_t start_byte,
                        SearchUnit unit) {
  const std::string lowered_haystack = to_lower_ascii(haystack);
  const std::string lowered_needle = to_lower_ascii(needle);
  // Fast path: a pattern with no `*`, `?` or `~` is a plain substring search.
  // A bare `~` still needs the wildcard path because `~?` / `~*` unescape.
  if (lowered_needle.find_first_of("*?~") == std::string::npos) {
    return lowered_haystack.find(lowered_needle, start_byte);
  }
  // The wildcard matchers report offsets relative to the scanned suffix.
  const std::string_view suffix = std::string_view(lowered_haystack).substr(start_byte);
  const bool sbcs_question =
      unit == SearchUnit::DbcsByte && question_spans_sbcs_only(locale_facts(current_eval_profile()).dbcs_codepage);
  const std::size_t rel =
      sbcs_question ? wildcard_find_dbcs(lowered_needle, suffix) : wildcard_find(lowered_needle, suffix);
  if (rel == std::string_view::npos) {
    return std::string::npos;
  }
  return start_byte + rel;
}

Value find_utf16_exact(const SearchArgs& args) {
  const std::size_t start_byte = utf16_to_byte_offset(args.haystack, static_cast<std::uint32_t>(args.start - 1));
  const std::size_t pos = args.haystack.find(args.needle, start_byte);
  if (pos == std::string::npos) {
    return Value::error(ErrorCode::Value);
  }
  const std::uint32_t units = utf16_units_in(std::string_view(args.haystack).substr(0, pos));
  return Value::number(static_cast<double>(units + 1));
}

Value find_utf16_folded(const SearchArgs& args) {
  const std::size_t start_byte = utf16_to_byte_offset(args.haystack, static_cast<std::uint32_t>(args.start - 1));
  const std::size_t pos = find_folded(args.haystack, args.needle, start_byte, SearchUnit::Utf16);
  if (pos == std::string::npos) {
    return Value::error(ErrorCode::Value);
  }
  const std::uint32_t units = utf16_units_in(std::string_view(args.haystack).substr(0, pos));
  return Value::number(static_cast<double>(units + 1));
}

Expected<TextWindowArgs, ErrorCode> read_text_window_args(const Value* args, std::uint32_t arity, bool snap_start) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return std::move(text.error());
  }
  auto start = snap_start ? read_snapped_int_arg(args[1], 1.0) : read_int_arg(args[1]);
  if (!start) {
    return std::move(start.error());
  }
  auto count = read_int_arg(args[2]);
  if (!count) {
    return std::move(count.error());
  }
  std::string new_text;
  if (arity > 3) {
    auto coerced = coerce_to_text(args[3]);
    if (!coerced) {
      return std::move(coerced.error());
    }
    new_text = std::move(coerced.value());
  }
  if (start.value() < 1 || count.value() < 0) {
    return ErrorCode::Value;
  }
  return TextWindowArgs{std::move(text.value()), start.value(), count.value(), std::move(new_text)};
}

}  // namespace text_detail
}  // namespace eval
}  // namespace formulon
