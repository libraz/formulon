// Normalises stored formula text into the bare expression the parser expects.
//
// Stored formula text may or may not carry a leading `=` depending on who
// wrote it (entry, the OOXML / XLSB readers, the C API). Every path that
// re-parses stored text strips it through this one helper so the readers,
// writers and evaluator cannot drift.

#ifndef FORMULON_UTILS_FORMULA_TEXT_H_
#define FORMULON_UTILS_FORMULA_TEXT_H_

#include <string_view>

namespace formulon {

/// Returns `s` with a single leading '=' removed, if present.
///
/// One-shot by design: `==1` is a parse error in Excel too, so it must reach
/// the parser as `=1` rather than be coerced. An empty string is returned
/// unchanged.
inline std::string_view strip_formula_prefix(std::string_view s) noexcept {
  if (!s.empty() && s.front() == '=') {
    s.remove_prefix(1);
  }
  return s;
}

}  // namespace formulon

#endif  // FORMULON_UTILS_FORMULA_TEXT_H_
