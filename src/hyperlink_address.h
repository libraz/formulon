#ifndef FORMULON_HYPERLINK_ADDRESS_H_
#define FORMULON_HYPERLINK_ADDRESS_H_

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
///   `file:////server/share/...`.
/// - Anything else is returned verbatim.
NormalizedHyperlinkTarget normalize_hyperlink_target(std::string_view input);

/// Stores `link.target` in normalised form and moves a split-off fragment into
/// `link.location` when none was given. Returns the text Excel writes into an
/// empty cell for the target: `<target> - <fragment>` when a fragment was
/// split off, the stored target otherwise.
std::string normalize_hyperlink(Hyperlink& link);

}  // namespace formulon

#endif  // FORMULON_HYPERLINK_ADDRESS_H_
