#ifndef FORMULON_HYPERLINK_ADDRESS_H_
#define FORMULON_HYPERLINK_ADDRESS_H_

#include <optional>
#include <string>
#include <string_view>

#include "sheet_metadata.h"

namespace formulon {

/// A hyperlink target in the form Excel stores it, plus the `#fragment` split
/// off an http/https/ftp URL.
struct NormalizedHyperlinkTarget {
  std::string target;
  std::string fragment;
};

/// Rewrites a hyperlink target the way Excel's Hyperlinks.Add stores it.
///
/// - http / https / ftp without userinfo: scheme and host lowercased, an empty
///   path becomes `/`, percent-escapes of unreserved characters decoded, and a
///   `#fragment` split into `fragment`. With userinfo the text is unchanged.
/// - mailto: only the scheme is lowercased.
/// - An absolute POSIX path becomes `file:///...`; a UNC path becomes
///   `file:////server/share/...`; a drive path (`C:\x`, `C:/x`) becomes
///   `file:///C:/x`. A drive letter is recognised before any scheme. A
///   `#fragment` is split off a local path too. A POSIX path holding `%`
///   stays a plain path, with only `%20` decoded to a space.
/// - Anything else is returned verbatim.
NormalizedHyperlinkTarget normalize_hyperlink_target(std::string_view input);

/// Stores `link.target` in normalised form and moves a split-off fragment into
/// `link.location` when none was given. Returns the text Excel writes into an
/// empty cell for the target: `<target> - <fragment>` when a fragment was
/// split off, the stored target otherwise.
std::string normalize_hyperlink(Hyperlink& link);

/// `abs_target`, a local absolute target (`file:///a/b`, `file:///C:/a`,
/// `file:////server/share/a`, or the same path without the scheme), relative
/// to a book in the canonical `directory` (`WorkbookLocation`), as a save
/// writes it: `../` per folder left, then the rest of the target. Empty when
/// `directory` is empty, the target is not local and absolute, the roots
/// differ in kind, the drive letters or UNC server/share differ (ASCII
/// case-insensitive), or no folder segment is shared; segments compare
/// case-sensitively. `%20` in a `file:` URL reads as a space; a stray `%` in
/// a plain path is literal and written `%25`.
std::optional<std::string> book_relative_target(std::string_view abs_target, std::string_view directory);

/// `target`, a relative local path (no scheme, root or drive), resolved
/// lexically against the canonical `directory` into an absolute `file:` URL.
/// `%20` reads as a space, `.` and empty segments drop, and `\` also
/// separates under a drive or UNC directory. Empty when `directory` is empty,
/// the target is not a relative local path, or `..` climbs above the root.
std::optional<std::string> resolve_relative_target(std::string_view target, std::string_view directory);

/// The rel target and `display` a save writes for `link` with the book in
/// the canonical `directory` (empty when none): the relative form when
/// `book_relative_target` gives one, with that text (plus ` - <location>`
/// when the link has a location) as the display unless the link has its own;
/// otherwise the link's own target and display. A stray `%` (one not starting
/// a `%XX` escape) in a plain POSIX path is written `%25`, as Excel saves it.
struct SavedHyperlink {
  std::string target;
  std::string display;
};
SavedHyperlink saved_hyperlink(const Hyperlink& link, std::string_view directory);

/// A hyperlink rel target as a load stores it: a plain POSIX path has its
/// `%20` read as a space, as Excel opens it; every other `%XX` and every other
/// target stays verbatim.
std::string loaded_hyperlink_target(std::string_view target);

}  // namespace formulon

#endif  // FORMULON_HYPERLINK_ADDRESS_H_
