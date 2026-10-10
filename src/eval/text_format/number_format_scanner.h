//
// Internal header -- do not include outside `src/eval/text_format/`.
//
// Scanning primitives shared between the format-string tokenizer and the
// section splitter / classifier. All helpers are pure / `noexcept` and
// operate on raw `std::string_view` slices of the user-supplied format.
//
// The token / section types these helpers feed are declared in
// `number_format_types.h`; the consumer translation units are
// `number_format_tokenizer.cpp` and `number_format_section.cpp`.

#ifndef FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_SCANNER_H_
#define FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_SCANNER_H_

#include <cstddef>
#include <string>
#include <string_view>

#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {

// Normalises the syntax-bearing full-width forms accepted by the TEXT()
// parser. This is deliberately not a general NFKC pass: quoted text,
// escape/underscore/asterisk payloads, and malformed UTF-8 are copied byte
// for byte, and so is a full-width form of a `literal_glyphs` character
// outside brackets. The returned string owns the normalised view consumed by
// the tokenizer, so literal token offsets remain stable for the whole parse.
std::string normalize_fullwidth_format_syntax(std::string_view fmt, std::string_view literal_glyphs);

// Returns the byte width of the UTF-8 scalar beginning at `i`, or one byte for
// malformed/truncated input. Callers use this for payload consumption so a
// multibyte literal is never split into independent escape tokens.
std::size_t utf8_scalar_width(std::string_view fmt, std::size_t i) noexcept;

// Returns true iff `c` is one of `yYmMdDhHsS`. Used to detect date-family
// tokens; AM/PM handled separately because they span multiple characters.
bool is_date_letter(char c) noexcept;

// Parses a run of `letter` characters (case-insensitive for the given
// target). Advances `*i` past the run. Returns the run length (>= 1 given
// the caller already matched at least one character).
std::size_t scan_run(std::string_view fmt, std::size_t& i, char letter) noexcept;

// Case-sensitive form of `scan_run`.
std::size_t scan_run_exact(std::string_view fmt, std::size_t& i, char letter) noexcept;

// Stored spelling of the indexed `ColorN` colour form.
inline constexpr std::string_view kStoredColorIndexPrefix = "Color";

// Returns true if `body` is a stored-spelling color qualifier: one of the
// English `Black`, `Blue`, `Cyan`, `Green`, `Magenta`, `Red`, `White`,
// `Yellow`, or `ColorN` with N in 1..56, ASCII case-insensitively. A
// localized name reaches here already rewritten by `localize_format`.
//
// The name is a prefix: anything after it in the same bracket is ignored, so
// `[Red1]` and `[Color1x]` are both accepted. Leading whitespace is not.
// The colour itself is discarded, so a match is inert like `[$...]`.
bool is_color_specifier(std::string_view body) noexcept;

// Detects an Excel conditional-section directive of the form `[op N]`,
// where `op` is one of `>`, `>=`, `<`, `<=`, `=`, `<>` and `N` is a literal
// double. Leading whitespace is tolerated (Excel accepts `[> 100]`).
//
// Returns:
//   * 1 on a successful parse: writes the operator and value to `*out_op` /
//     `*out_value`.
//   * -1 if the body looks like a predicate (starts with one of the six
//     operator forms) but the numeric tail fails to parse fully. The caller
//     should treat this as an invalid bracket so `apply_format` surfaces
//     `#VALUE!` exactly as the legacy fallback path did.
//   * 0 if the body does not look like a predicate at all (caller should
//     continue trying other bracket interpretations).
int parse_cond_directive(std::string_view body, CondOp* out_op, double* out_value) noexcept;

// Detects `[DBNumN]` (body length exactly 6 bytes after stripping brackets).
// Matching is case-insensitive: `[DBNum1]`, `[dbnum2]`, `[DbNum3]` all parse.
// Returns the directive index (1 to 4) on a hit, otherwise 0.
int parse_dbnum_directive(std::string_view body) noexcept;

// Reads the spec after the `-` of a `[$symbol-spec]` tag into `tag`: hex
// `NNCCLLLL` (numeral system, calendar, language), or a name such as `ja-JP`,
// `ja-JP-x-gannen` or `x-sysdate` with an optional `,NNCC` hex suffix. An
// unknown name leaves `tag` unchanged; false for a spec Excel rejects.
bool parse_format_tag(std::string_view spec, FormatTag* tag) noexcept;

// Returns true if `tok` is a date-family token (including elapsed brackets).
bool is_date_tok(Tok t) noexcept;

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_SCANNER_H_
