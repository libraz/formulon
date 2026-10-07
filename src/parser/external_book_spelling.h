//
// Rewrites the `[N]` link indexes of a stored formula into the book names a
// formula bar shows, and maps a link target to the directory it displays.
//
// The parser stays free of workbook types: the link table is reached through
// `ExternalBookResolver`.

#ifndef FORMULON_PARSER_EXTERNAL_BOOK_SPELLING_H_
#define FORMULON_PARSER_EXTERNAL_BOOK_SPELLING_H_

#include <cstdint>
#include <string>
#include <string_view>

namespace formulon {
namespace parser {

/// Display parts of one external link: `path` is the directory prefix with
/// its trailing separator (empty when the link has none) and `book` the file
/// name.
struct ExternalBookDisplay {
  std::string path;
  std::string book;
};

/// Maps a 1-based link index to its display parts. `resolve` returns false
/// when no such link exists.
struct ExternalBookResolver {
  bool (*resolve)(const void* ctx, std::uint32_t index, ExternalBookDisplay* out);
  const void* ctx;
};

/// Replaces every `[N]Sheet!`, `[N]S1:S2!`, `'[N]...'!` and book-scope
/// `[N]!Name` qualifier in `stored` with its formula-bar spelling, quoted by
/// the external-qualifier rules. An index without a link is spelled as the
/// decimal book name. Bytes outside the rewritten qualifiers are kept, and
/// `[0]!Name`, structured references and string literals are left alone.
std::string spell_external_books(std::string_view stored, const ExternalBookResolver& resolver);

/// Decodes every `%XX` escape in `s`; a malformed escape is kept as written.
std::string percent_decode(std::string_view s);

/// The directory a link target displays as a qualifier path, with its
/// trailing separator, or empty when the target has none (a relative path or
/// a bare file name). `file:` URLs are percent-decoded; a Windows drive or
/// UNC target uses backslashes.
std::string display_path_for_link_target(std::string_view target);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_EXTERNAL_BOOK_SPELLING_H_
