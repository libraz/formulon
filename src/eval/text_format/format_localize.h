// Maps a localized TEXT format string onto the invariant format syntax.
//
// Internal header -- do not include outside `src/eval/text_format/`.
//
// A TEXT() argument is spelled in the UI locale: its decimal and group
// separators, colour names and General keyword are the locale's own. This
// pre-pass rewrites those into the invariant spelling the tokenizer reads
// (`.`, `,`, `[Red]`, `General`) and folds the ja-JP full-width syntax. Date
// letters are not rewritten here; the tokenizer reads them through the
// locale's `FormatLetters`. Quoted text, escape payloads and `_X` / `*X`
// payloads are copied byte for byte.

#ifndef FORMULON_EVAL_TEXT_FORMAT_FORMAT_LOCALIZE_H_
#define FORMULON_EVAL_TEXT_FORMAT_FORMAT_LOCALIZE_H_

#include <string>
#include <string_view>

#include "eval/text_format/number_format.h"
#include "excel_profile.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {

/// A format code in the invariant syntax.
struct LocalizedFormat {
  std::string text;
  /// False when the code uses a spelling the locale rejects, such as the
  /// English `[Red]` or `General` outside en-US.
  bool valid = true;
};

/// Rewrites `fmt` for the tokenizer. A stored code (`kStored`) is already
/// invariant and only takes the full-width fold of a folding locale.
LocalizedFormat localize_format(std::string_view fmt, FormatDialect dialect, ExcelProfile profile);

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_FORMAT_LOCALIZE_H_
